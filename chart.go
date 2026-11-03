// Copyright 2026 kenny-not-dead. All rights reserved.
// Use of this source code is governed by a BSD-style license
// that can be found in the LICENSE file.
//
// Package gopptx provides functionality to create and manipulate PowerPoint
// (.pptx) files in Go, using the Office Open XML (ECMA-376) format.

package gopptx

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
)

// the ids of the axes of the generated chart: they only have to be unique
// within the chart part. The ids are positive, the same way PowerPoint
// writes them: the val attribute of the c:axId and the c:crossAx elements
// is an unsigned int.
const (
	chartCatAxisID = 2068027336
	chartValAxisID = 2113994440
)

// externalDataRe matches the reference to the external data of the chart
// (c:externalData, the embedded workbook) in the chart part XML.
var externalDataRe = regexp.MustCompile(`(?s)<[A-Za-z_][\w.-]*:externalData\b[^>]*/>|<[A-Za-z_][\w.-]*:externalData\b[^>]*>.*?</[A-Za-z_][\w.-]*:externalData>`)

// ChartOptions can be passed to AddChart and AddChartFromBytes to set the
// title, the name, the position and the size of the new chart. By default
// (the zero values of the position and the size) the chart is centered on
// the slide, taking 80% of the slide size, the same way the tables do it.
//
// The chart always has the legend at the bottom: when the chart has the
// title, the height of the frame should be about 2 500 000 EMU (about
// 7 cm) or more, otherwise the text rows of the chart (the title, the
// category labels and the legend) can overlap in PowerPoint.
type ChartOptions struct {
	// Title is the title of the chart, by default the chart has no title.
	Title string
	// Name is the name of the graphic frame (the name attribute of the
	// p:cNvPr element), by default "Chart <id>".
	Name string
	// OffsetX, OffsetY are the x and the y offsets of the chart in EMU.
	OffsetX int
	OffsetY int
	// Width, Height are the width and the height of the chart in EMU.
	Width  int
	Height int
}

// ChartSeries is a series of the new chart: the name of the series and the
// values of the points, one value per category.
type ChartSeries struct {
	Name   string
	Values []float64
}

// AddChart provides a function to add a new chart to the slide by the given
// slide id and returns the id of the chart in the slide after it is
// appended. The chart is a column chart (c:barChart with the barDir "col")
// with the given categories and the series: the data of the chart is stored
// in the cache of the chart part (ppt/charts/chartN.xml), the chart part and
// its content type are created automatically, and the chart relationship is
// added to the slide rels part. The new graphic frame is appended to the end
// of the shape tree: it is drawn above the existing shapes of the slide.
//
//	chartID, err := f.AddChart(slideID,
//		[]string{"2024", "2025", "2026"},
//		[]gopptx.ChartSeries{{Name: "Revenue", Values: []float64{120, 180, 250}}},
//		&gopptx.ChartOptions{Title: "Revenue"})
func (f *File) AddChart(slideID int, categories []string, series []ChartSeries, opts *ChartOptions) (int, error) {
	if len(categories) == 0 || len(series) == 0 {
		return -1, ErrChartInvalidData
	}
	for _, s := range series {
		if len(s.Values) != len(categories) {
			return -1, ErrChartInvalidData
		}
	}
	slide, err := f.slideReader(slideID)
	if err != nil {
		return -1, err
	}

	var title string
	if opts != nil {
		title = opts.Title
	}
	data := newChartSpaceXML(categories, series, title)

	return f.addChartPart(slideID, slide, data, opts)
}

// AddChartFromBytes provides a function to add a new chart to the slide by
// the given slide id and the data of the chart part (the c:chartSpace
// element of the ppt/charts/chartN.xml part) and returns the id of the chart
// in the slide after it is appended. The data of the chart part must be
// self-contained (the data of the series in the cache): the reference to the
// external data (c:externalData, the embedded workbook) is removed from the
// given data. The chart part, its content type and the chart relationship
// in the slide rels part are created automatically, the same way AddChart
// does it.
func (f *File) AddChartFromBytes(slideID int, data []byte, opts *ChartOptions) (int, error) {
	if len(data) == 0 {
		return -1, ErrChartEmpty
	}
	slide, err := f.slideReader(slideID)
	if err != nil {
		return -1, err
	}

	data = externalDataRe.ReplaceAll(data, nil)
	if !bytes.HasPrefix(data, []byte("<")) {
		data = []byte(xml.Header + string(data))
	}

	return f.addChartPart(slideID, slide, data, opts)
}

// GetCharts provides a function to get the charts of the slide by the given
// slide id, in the document order of the shape tree. The charts are the
// graphic frames (p:graphicFrame) with the chart reference (c:chart),
// returned by value: edits of the returned slice do not affect the slide.
func (f *File) GetCharts(slideID int) ([]decodeGraphicFrame, error) {
	var charts []decodeGraphicFrame
	slide, err := f.slideReader(slideID)
	if err != nil {
		return charts, err
	}

	for i := range slide.CommonSlideData.ShapeTree.Children {
		child := &slide.CommonSlideData.ShapeTree.Children[i]
		if child.GraphicFrame == nil {
			continue
		}
		if _, ok := child.GraphicFrame.chartRef(); !ok {
			continue
		}
		charts = append(charts, *child.GraphicFrame)
	}

	return charts, nil
}

// GetChart provides a function to get the chart of the slide by the given
// slide id and chart id (the id attribute of the p:cNvPr element of the
// graphic frame that contains the chart reference), the same way GetShapes
// returns the shapes: the graphic frame is returned by value.
func (f *File) GetChart(slideID, chartID int) (decodeGraphicFrame, error) {
	gf, err := f.chartGraphicFrame(slideID, chartID)
	if err != nil {
		return decodeGraphicFrame{}, err
	}

	return *gf, nil
}

// DeleteChart provides a function to delete the chart from the slide by the
// given slide id and chart id. The graphic frame is removed from the shape
// tree of the slide, the chart relationship is removed from the slide rels
// part, and the chart part (and the parts that it references, for example
// the embedded workbook) is removed from the package when no other part of
// the presentation references it anymore. The content type of the chart part
// is removed from the content types part.
func (f *File) DeleteChart(slideID int, chartID int) error {
	slide, err := f.slideReader(slideID)
	if err != nil {
		return err
	}

	children := slide.CommonSlideData.ShapeTree.Children
	var gf *decodeGraphicFrame
	deleteIndex := -1
	for i := range children {
		child := &children[i]
		if child.GraphicFrame == nil ||
			child.GraphicFrame.NonVisualGraphicFrameProperties == nil ||
			child.GraphicFrame.NonVisualGraphicFrameProperties.CommonNonVisualProperties == nil {
			continue
		}
		if child.GraphicFrame.NonVisualGraphicFrameProperties.CommonNonVisualProperties.ID == chartID {
			if _, ok := child.GraphicFrame.chartRef(); !ok {
				return ErrChartNotExist{chartID}
			}
			gf = child.GraphicFrame
			deleteIndex = i
			break
		}
	}
	if deleteIndex == -1 {
		return ErrChartNotExist{chartID}
	}

	slide.CommonSlideData.ShapeTree.Children = append(children[:deleteIndex], children[deleteIndex+1:]...)

	// the chart part path by the chart relationship of the slide rels part
	slidePath, _ := f.getSlideXMLPath(slideID)
	chartPath := f.chartPartPath(slidePath, gf)
	if chartPath == "" {
		return nil
	}

	// remove the chart relationship from the slide rels part
	rels, err := f.relsReader(f.getSlideRelsPath(slidePath))
	if err != nil {
		return err
	}
	if rels != nil {
		rels.mu.Lock()
		for i, rel := range rels.Relationships {
			if rel.Type == SourceRelationshipChart && resolvePartPath(slidePath, rel.Target) == chartPath {
				rels.Relationships = append(rels.Relationships[:i], rels.Relationships[i+1:]...)
				break
			}
		}
		rels.mu.Unlock()
	}

	// remove the chart part, its rels part and the parts that it references
	// (for example the embedded workbook) when no other part of the
	// presentation references them anymore
	if !f.partReferenced(chartPath, SourceRelationshipChart) {
		f.Pkg.Delete(chartPath)
		f.removeContentTypeIfUnused(chartPath)
		f.deleteChartRelsAndTargets(chartPath)
	}

	return nil
}

// addChartPart provides a function to create the chart part in the package,
// the chart relationship in the slide rels part and the graphic frame (the
// c:chart reference) in the shape tree of the slide, and returns the id of
// the chart in the slide.
func (f *File) addChartPart(slideID int, slide *decodeSlide, data []byte, opts *ChartOptions) (int, error) {
	var o ChartOptions
	if opts != nil {
		o = *opts
	}

	partName := f.newChartPartName()
	f.Pkg.Store(partName, data)
	_ = f.setContentTypes("/"+partName, ContentTypeChartML)

	slidePath, _ := f.getSlideXMLPath(slideID)
	// the chart reference (c:chart) uses the r prefix, so the slide root
	// element must declare the relationships namespace
	f.addNameSpaces(slidePath, SourceRelationship)
	target := "../" + strings.TrimPrefix(partName, "ppt/")
	rID := f.addRels(f.getSlideRelsPath(slidePath), SourceRelationshipChart, target, "")

	chartID := f.nextShapeID(slide)
	offset, extents := f.frameLayout(o.OffsetX, o.OffsetY, o.Width, o.Height)
	name := o.Name
	if name == "" {
		name = fmt.Sprintf("Chart %d", chartID)
	}

	graphicFrame := decodeGraphicFrame{
		NonVisualGraphicFrameProperties: &decodeNonVisualGraphicFrameProperties{
			CommonNonVisualProperties: &decodeCommonNonVisualProperties{
				ID:   chartID,
				Name: name,
			},
			// the graphic frame locks element and the non-visual properties
			// element are written the same way PowerPoint writes them
			CommonNonVisualGraphicFrameProperties: &innerXML{Content: `<a:graphicFrameLocks noGrp="1"/>`},
			NonVisualProperties:                   &innerXML{},
		},
		Xfrm: &DecodeXfrm{
			Offset:  &offset,
			Extents: &extents,
		},
		Graphic: &decodeGraphic{
			GraphicData: &decodeGraphicData{
				URI:   graphicDataURIChart,
				Chart: &decodeChart{ID: "rId" + strconv.Itoa(rID)},
			},
		},
	}

	// the new graphic frame is appended to the end of the shape tree: it is
	// drawn above the existing shapes of the slide
	slide.CommonSlideData.ShapeTree.Children = append(
		slide.CommonSlideData.ShapeTree.Children,
		decodeShapeTreeChild{GraphicFrame: &graphicFrame},
	)

	return chartID, nil
}

// newChartPartName provides a function to generate the next free chart part
// name (ppt/charts/chartN.xml) in the presentation.
func (f *File) newChartPartName() string {
	maxIdx := 0
	findMax := func(name string) {
		if strings.HasPrefix(name, "ppt/charts/chart") && strings.HasSuffix(name, ".xml") {
			base := strings.TrimSuffix(strings.TrimPrefix(name, "ppt/charts/chart"), ".xml")
			if n, err := strconv.Atoi(base); err == nil && n > maxIdx {
				maxIdx = n
			}
		}
	}

	f.Pkg.Range(func(k, v interface{}) bool {
		name, _ := k.(string)
		findMax(name)
		return true
	})

	f.tempFiles.Range(func(k, v interface{}) bool {
		name, _ := k.(string)
		findMax(name)
		return true
	})

	return "ppt/charts/chart" + strconv.Itoa(maxIdx+1) + ".xml"
}

// chartPartPath provides a function to get the path of the chart part by the
// given slide part path and the graphic frame that contains the chart
// reference: the target of the chart relationship in the slide rels part.
func (f *File) chartPartPath(slidePath string, gf *decodeGraphicFrame) string {
	ref, _ := gf.chartRef()
	rels, _ := f.relsReader(f.getSlideRelsPath(slidePath))
	if rels == nil {
		return ""
	}
	rels.mu.Lock()
	defer rels.mu.Unlock()
	for _, rel := range rels.Relationships {
		if rel.ID == ref.ID && rel.Type == SourceRelationshipChart {
			return resolvePartPath(slidePath, rel.Target)
		}
	}

	return ""
}

// deleteChartRelsAndTargets provides a function to delete the rels part of
// the chart part and the parts that it references (for example the embedded
// workbook) when no other part of the presentation references them anymore.
func (f *File) deleteChartRelsAndTargets(chartPath string) {
	relsPath := chartRelsPath(chartPath)
	targets := f.relsTargets(relsPath)
	f.Pkg.Delete(relsPath)
	// the rels part is loaded to the Relationships map by the relsTargets
	// function: it must be removed from the map too, otherwise the
	// relsWriter writes it back on save
	f.Relationships.Delete(relsPath)
	for _, target := range targets {
		if f.partReferenced(target, "") {
			continue
		}
		f.Pkg.Delete(target)
		f.removeContentTypeIfUnused(target)
	}
}

// chartRelsPath provides a function to get the path of the rels part of the
// chart part.
func chartRelsPath(chartPath string) string {
	return filepath.ToSlash(filepath.Join(filepath.Dir(chartPath), "_rels", filepath.Base(chartPath)+".rels"))
}

// relsTargets provides a function to get the resolved targets of the
// relationships of the given rels part.
func (f *File) relsTargets(relsPath string) []string {
	rels, _ := f.relsReader(relsPath)
	if rels == nil {
		return nil
	}
	rels.mu.Lock()
	defer rels.mu.Unlock()
	var targets []string
	base := basePartOfRelsPart(relsPath)
	for _, rel := range rels.Relationships {
		targets = append(targets, resolvePartPath(base, rel.Target))
	}
	return targets
}

// chartGraphicFrame provides a function to get the graphic frame that
// contains the chart reference by the given slide id and chart id (the id
// attribute of the p:cNvPr element of the graphic frame).
func (f *File) chartGraphicFrame(slideID, chartID int) (*decodeGraphicFrame, error) {
	slide, err := f.slideReader(slideID)
	if err != nil {
		return nil, err
	}

	for i := range slide.CommonSlideData.ShapeTree.Children {
		child := &slide.CommonSlideData.ShapeTree.Children[i]
		if child.GraphicFrame == nil ||
			child.GraphicFrame.NonVisualGraphicFrameProperties == nil ||
			child.GraphicFrame.NonVisualGraphicFrameProperties.CommonNonVisualProperties == nil {
			continue
		}
		if child.GraphicFrame.NonVisualGraphicFrameProperties.CommonNonVisualProperties.ID == chartID {
			if _, ok := child.GraphicFrame.chartRef(); !ok {
				return nil, ErrChartNotExist{chartID}
			}
			return child.GraphicFrame, nil
		}
	}

	return nil, ErrChartNotExist{chartID}
}

// frameLayout provides a function to calculate the offset and the extents
// (in EMU) of a new graphic frame (the table or the chart). By default (the
// zero values of the position and the size) the frame is centered on the
// slide, taking 80% of the slide size, the same way pictureLayout does it
// for the pictures.
func (f *File) frameLayout(offsetX, offsetY, width, height int) (Offset, Extents) {
	slideSize := slideSize{CY: 5143500, CX: 9144000}
	if presentation, _ := f.presentationReader(); presentation != nil && presentation.SlideSize != nil {
		slideSize = *presentation.SlideSize
	}

	offset := Offset{X: slideSize.CX / 10, Y: slideSize.CY / 10}
	extents := Extents{CX: slideSize.CX * 4 / 5, CY: slideSize.CY * 4 / 5}
	if offsetX != 0 {
		offset.X = offsetX
	}
	if offsetY != 0 {
		offset.Y = offsetY
	}
	if width != 0 {
		extents.CX = width
	}
	if height != 0 {
		extents.CY = height
	}

	return offset, extents
}

// newChartSpaceXML provides a function to generate the data of the chart
// part (the c:chartSpace element): a column chart (c:barChart with the
// barDir "col") with the given categories and the series, the data of the
// chart is stored in the cache of the series (c:strCache and c:numCache),
// the chart has no external data (no embedded workbook). The structure of
// the element mirrors the data that PowerPoint writes for the same chart:
// the language and the rounded corners of the chart space, the explicit
// text properties of the chart (the default font size of the text that has
// no size of its own), the numeric format and the text properties of the
// axes, and the text properties of the legend.
func newChartSpaceXML(categories []string, series []ChartSeries, title string) []byte {
	esc := func(s string) string {
		var buf bytes.Buffer
		_ = xml.EscapeText(&buf, []byte(s))
		return buf.String()
	}
	num := func(v float64) string {
		return strconv.FormatFloat(v, 'f', -1, 64)
	}
	col := func(i int) string {
		return string(rune('A' + i + 1))
	}

	// the text properties (c:txPr) of the axes and the legend: the text of
	// the chart has the default size of 12 pt, the same way the chart
	// created by PowerPoint does it
	const txPr12 = `<a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1200"/></a:pPr><a:endParaRPr lang="en-US"/></a:p>`

	var b strings.Builder
	b.WriteString(`<c:chartSpace xmlns:c="` + NameSpaceDrawingMLChart.Value + `" xmlns:a="` + NameSpaceDrawingMLMain + `">`)
	b.WriteString(`<c:date1904 val="0"/>`)
	b.WriteString(`<c:lang val="en-US"/>`)
	b.WriteString(`<c:roundedCorners val="0"/>`)
	b.WriteString(`<c:chart>`)
	if title != "" {
		b.WriteString(`<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="1800"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="75000"/></a:schemeClr></a:solidFill><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mj-cs"/></a:rPr><a:t>`)
		b.WriteString(esc(title))
		b.WriteString(`</a:t></a:r></a:p></c:rich></c:tx><c:layout/><c:overlay val="0"/></c:title>`)
		b.WriteString(`<c:autoTitleDeleted val="0"/>`)
	} else {
		b.WriteString(`<c:autoTitleDeleted val="1"/>`)
	}
	b.WriteString(`<c:plotArea><c:layout/>`)
	b.WriteString(`<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>`)
	for i, s := range series {
		b.WriteString(`<c:ser><c:idx val="` + strconv.Itoa(i) + `"/><c:order val="` + strconv.Itoa(i) + `"/>`)
		b.WriteString(`<c:tx><c:strRef><c:f>Sheet1!$` + col(i) + `$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>`)
		b.WriteString(esc(s.Name))
		b.WriteString(`</c:v></c:pt></c:strCache></c:strRef></c:tx>`)
		b.WriteString(`<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$` + strconv.Itoa(len(categories)+1) + `</c:f><c:strCache><c:ptCount val="` + strconv.Itoa(len(categories)) + `"/>`)
		for j, c := range categories {
			b.WriteString(`<c:pt idx="` + strconv.Itoa(j) + `"><c:v>`)
			b.WriteString(esc(c))
			b.WriteString(`</c:v></c:pt>`)
		}
		b.WriteString(`</c:strCache></c:strRef></c:cat>`)
		b.WriteString(`<c:val><c:numRef><c:f>Sheet1!$` + col(i) + `$2:$` + col(i) + `$` + strconv.Itoa(len(categories)+1) + `</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="` + strconv.Itoa(len(categories)) + `"/>`)
		for j, v := range s.Values {
			b.WriteString(`<c:pt idx="` + strconv.Itoa(j) + `"><c:v>` + num(v) + `</c:v></c:pt>`)
		}
		b.WriteString(`</c:numCache></c:numRef></c:val>`)
		b.WriteString(`</c:ser>`)
	}
	b.WriteString(`<c:gapWidth val="150"/><c:overlap val="0"/>`)
	b.WriteString(`<c:axId val="` + strconv.Itoa(chartCatAxisID) + `"/><c:axId val="` + strconv.Itoa(chartValAxisID) + `"/>`)
	b.WriteString(`</c:barChart>`)
	b.WriteString(`<c:catAx><c:axId val="` + strconv.Itoa(chartCatAxisID) + `"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:txPr>` + txPr12 + `</c:txPr><c:crossAx val="` + strconv.Itoa(chartValAxisID) + `"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`)
	b.WriteString(`<c:valAx><c:axId val="` + strconv.Itoa(chartValAxisID) + `"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:txPr>` + txPr12 + `</c:txPr><c:crossAx val="` + strconv.Itoa(chartCatAxisID) + `"/><c:crosses val="autoZero"/></c:valAx>`)
	b.WriteString(`</c:plotArea>`)
	b.WriteString(`<c:legend><c:legendPos val="b"/><c:layout/><c:txPr>` + txPr12 + `</c:txPr><c:overlay val="0"/></c:legend>`)
	b.WriteString(`<c:plotVisOnly val="1"/>`)
	b.WriteString(`<c:dispBlanksAs val="gap"/>`)
	b.WriteString(`<c:showDLblsOverMax val="0"/>`)
	b.WriteString(`</c:chart>`)
	// the default text properties of the chart: the base font size of the
	// text that has no size of its own (without the element the text falls
	// back to the bigger default size of the theme)
	b.WriteString(`<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1800"/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>`)
	b.WriteString(`</c:chartSpace>`)

	return []byte(xml.Header + b.String())
}

// allRels provides a function to get the relationships of all the parts of
// the presentation, by the part path: the loaded rels parts (the
// Relationships map) and the raw rels parts in the package.
func (f *File) allRels() map[string][]relationship {
	m := make(map[string][]relationship)
	f.Relationships.Range(func(k, v interface{}) bool {
		name, _ := k.(string)
		if r, ok := v.(*relationships); ok {
			m[basePartOfRelsPart(name)] = r.Relationships
		}
		return true
	})
	f.Pkg.Range(func(k, v interface{}) bool {
		name, _ := k.(string)
		if content, ok := v.([]byte); ok && strings.HasSuffix(name, ".rels") {
			base := basePartOfRelsPart(name)
			if _, ok := m[base]; ok {
				return true
			}
			rels := relationships{}
			if err := xml.Unmarshal(namespaceStrictToTransitional(content), &rels); err == nil {
				m[base] = rels.Relationships
			}
		}
		return true
	})

	return m
}

// partReferenced reports whether the given part is referenced by a
// relationship of the given type (of any type when relType is empty) of any
// part of the presentation.
func (f *File) partReferenced(partPath, relType string) bool {
	for base, rels := range f.allRels() {
		for _, rel := range rels {
			if relType != "" && rel.Type != relType {
				continue
			}
			if resolvePartPath(base, rel.Target) == partPath {
				return true
			}
		}
	}
	return false
}

// removeContentTypeIfUnused provides a function to remove the content type
// of the given part from the content types part when the part is removed
// from the package: the override of the part and the default of the part
// extension, when it is not used by any other part of the package anymore
// and is not declared in the origin content types part.
func (f *File) removeContentTypeIfUnused(partPath string) {
	content, err := f.contentTypesReader()
	if err != nil {
		return
	}
	content.mu.Lock()
	defer content.mu.Unlock()
	for i, o := range content.Overrides {
		if o.PartName == "/"+partPath {
			content.Overrides = append(content.Overrides[:i], content.Overrides[i+1:]...)
			break
		}
	}

	ext := strings.TrimPrefix(filepath.Ext(partPath), ".")
	if ext == "" {
		return
	}
	inUse := false
	f.Pkg.Range(func(k, v interface{}) bool {
		name, _ := k.(string)
		if strings.EqualFold(filepath.Ext(name), "."+ext) {
			inUse = true
			return false
		}
		return true
	})
	if inUse {
		return
	}
	
	for i, d := range content.Defaults {
		if !strings.EqualFold(d.Extension, ext) {
			continue
		}
		// keep the default declared in the origin content types part
		if inStrSlice(f.originalContentTypeExtensions, d.Extension, false) != -1 {
			return
		}
		content.Defaults = append(content.Defaults[:i], content.Defaults[i+1:]...)
		return
	}
}
