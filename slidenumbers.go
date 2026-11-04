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
	"html"
	"io"
	"path/filepath"
	"regexp"
	"sort"
	"strconv"
	"strings"
)

// SlideNumberOptions defines the appearance of the slide number placeholder
// (the sldNum placeholder of the slide layouts) set by SetSlideNumbers: the
// position and the size of the placeholder (Xfrm) and the properties of the
// number text (Run: the font, the size in hundredths of a point, the color,
// the bold, and so on). The zero values keep the current values of the
// layout (the standard position of the layout and the text properties of
// the field of the placeholder).
type SlideNumberOptions struct {
	// Xfrm is the position and the size of the slide number placeholder in
	// the layout: the offset from the top-left corner of the slide and the
	// extents, in EMU.
	Xfrm *DecodeXfrm
	// Run is the properties of the number text written to the run
	// properties (a:rPr) of the slide number field (a:fld) of the
	// placeholder: the size in hundredths of a point (Size), the color
	// (SolidFill), the typeface (Latin, EastAsian, ComplexScript), the bold
	// (Bold), and so on.
	Run *DecodeRunProperties
}

// SetSlideNumbers enables (or disables, when enable is false) the slide
// numbers of the presentation by the standard mechanism (the same as the
// "Insert > Slide number" dialog of PowerPoint): the "sldNum" flag of the
// "p:hf" element of each slide master of the presentation is set, the slide
// number placeholder (p:ph type="sldNum") with the slide number field
// (a:fld type="slidenum") is ensured in each slide layout, and the slide
// number placeholder is ensured on each slide of the presentation.
//
// A slide shows the number only when it has the slide number placeholder of
// its own and the "sldNum" flag of the slide master is enabled, so the
// placeholder is added to the slides that have no such placeholder, and the
// new slides of the presentation (created by NewSlide after the numbers are
// enabled) get the placeholder on the creation. The placeholder of the slide
// has the slide number field of the same id and the same cached value as the
// field of the placeholder of the layout of the slide (the same as the files
// of Google Slides), the position and the format of the number are inherited
// from the placeholder of the layout. When enable is false, the flag is set
// to "0" and the number is hidden (the placeholders are kept in the layouts
// and the slides).
//
// The position and the format of the number are defined in the slide
// layouts; the SlideNumberOptions are applied to the placeholder of each
// layout (the layout that has no the placeholder gets a new one, with the
// standard index 12).
//
// The text box with the slide number field on a single slide (the manual
// mechanism of NewSlideNumberField) is a custom alternative (the number on
// a part of the slides, or in a custom position of one slide only); it is
// not connected to the slide layouts, so the new slides have no such text
// box.
//
//	if err := f.SetSlideNumbers(true); err != nil {
//	    fmt.Println(err)
//	}
//
// the number in the bottom-right corner, 18 pt, in the Arial typeface:
//
//	size := 1800
//	if err := f.SetSlideNumbers(true, gopptx.SlideNumberOptions{
//	    Xfrm: &gopptx.DecodeXfrm{
//	        Offset:  &gopptx.Offset{X: 8709025, Y: 4756150},
//	        Extents: &gopptx.Extents{CX: 914400, CY: 457200},
//	    },
//	    Run: &gopptx.DecodeRunProperties{
//	        Size:  &size,
//	        Latin: &gopptx.Font{Typeface: "Arial"},
//	    },
//	}); err != nil {
//	    fmt.Println(err)
//	}
func (f *File) SetSlideNumbers(enable bool, opts ...SlideNumberOptions) error {
	options := SlideNumberOptions{}
	for _, o := range opts {
		if o.Xfrm != nil {
			options.Xfrm = o.Xfrm
		}
		if o.Run != nil {
			options.Run = o.Run
		}
	}

	for _, path := range f.slideMasterPaths() {
		if err := f.setMasterSlideNumberFlag(path, enable); err != nil {
			return err
		}
	}
	if !enable {
		return nil
	}

	for _, path := range f.slideLayoutPaths() {
		if err := f.ensureLayoutSlideNumber(path, options); err != nil {
			return err
		}
	}

	// the slides: the slide number placeholder of each slide (the empty
	// placeholder, the position and the format of the number are inherited
	// from the placeholder of the layout of the slide); the parts of the
	// slides that already have the placeholder are not modified
	if slides, err := f.getSlideMap(); err != nil {
		return err
	} else {
		for slideID := range slides {
			if err := f.ensureSlideSlideNumber(slideID); err != nil {
				return err
			}
		}
	}

	return nil
}

// GetSlideNumbers returns the state of the slide numbers of the
// presentation: the value of the "sldNum" flag of the "p:hf" element of the
// first slide master (the same state the "Insert > Slide number" dialog of
// PowerPoint shows). When the "p:hf" element or the "sldNum" attribute is
// absent, the default value of the flag by the schema (true) is returned.
// The flag controls the visibility of the slide number placeholder of the
// slide layouts; it does not report whether the layouts contain the
// placeholder.
func (f *File) GetSlideNumbers() bool {
	paths := f.slideMasterPaths()
	if len(paths) == 0 {
		return false
	}

	tag := masterHFElementRe.Find(f.readBytes(paths[0]))
	if tag == nil {
		return true
	}

	m := masterHFSldNumRe.FindSubmatch(tag)
	if m == nil {
		return true
	}

	v := string(m[1])

	return v != "0" && v != "false"
}

// slideMasterPaths returns the part paths of the slide masters of the
// presentation, sorted by name.
func (f *File) slideMasterPaths() []string {
	return f.partPathsByPrefix("ppt/slideMasters/slideMaster")
}

// slideLayoutPaths returns the part paths of the slide layouts of the
// presentation, sorted by name.
func (f *File) slideLayoutPaths() []string {
	return f.partPathsByPrefix("ppt/slideLayouts/slideLayout")
}

// partPathsByPrefix returns the part names of the package that start with
// the given prefix and end with ".xml" (the ".rels" parts are not included),
// sorted by name.
func (f *File) partPathsByPrefix(prefix string) []string {
	var paths []string
	f.Pkg.Range(func(k, _ interface{}) bool {
		name, _ := k.(string)
		if strings.HasPrefix(name, prefix) && strings.HasSuffix(name, ".xml") {
			paths = append(paths, name)
		}
		return true
	})
	sort.Strings(paths)
	return paths
}

// masterHFElementRe matches the "p:hf" element (the header footer element,
// CT_HeaderFooter) of the slide master.
var masterHFElementRe = regexp.MustCompile(`<(?:[A-Za-z0-9]+:)?hf(?:\s[^>]*)?/?>`)

// masterHFSldNumRe matches the "sldNum" attribute of the "p:hf" element.
var masterHFSldNumRe = regexp.MustCompile(`sldNum="([^"]*)"`)

// masterRootPrefixRe matches the namespace prefix of the root element of
// the slide master part.
var masterRootPrefixRe = regexp.MustCompile(`^<\?xml[^>]*\?>\s*<([A-Za-z0-9]+):sldMaster`)

// setMasterSlideNumberFlag provides a function to set the "sldNum" flag of
// the "p:hf" element of the slide master (creating the element when it is
// absent): the flag enables or disables the slide number placeholder of the
// presentation.
func (f *File) setMasterSlideNumberFlag(path string, enable bool) error {
	raw := f.readBytes(path)
	value := "0"
	if enable {
		value = "1"
	}

	tag := masterHFElementRe.Find(raw)
	if tag != nil {
		updated := string(tag)
		if masterHFSldNumRe.MatchString(updated) {
			updated = masterHFSldNumRe.ReplaceAllString(updated, `sldNum="`+value+`"`)
		} else if i := strings.IndexAny(updated, " />"); i != -1 {
			// the attribute is inserted after the element name (the
			// "sldNum" attribute is the first one of the element by the
			// schema)
			updated = updated[:i] + ` sldNum="` + value + `"` + updated[i:]
		}
		raw = bytes.Replace(raw, tag, []byte(updated), 1)
	} else {
		// the element is absent: it is inserted before the first of the
		// "txStyles" and the "extLst" elements (the order of the elements
		// of the slide master by the schema), or before the end of the
		// slide master root element
		prefix := "p:"
		if m := masterRootPrefixRe.FindSubmatch(raw); m != nil {
			prefix = string(m[1]) + ":"
		}
		insertion := "<" + prefix + `hf sldNum="` + value + `" />`
		pos := len(raw)
		for _, el := range []string{"txStyles", "extLst"} {
			if i := bytes.Index(raw, []byte("<"+prefix+el)); i != -1 && i < pos {
				pos = i
			}
		}
		if pos == len(raw) {
			if i := bytes.LastIndex(raw, []byte("</"+prefix+"sldMaster>")); i != -1 {
				pos = i
			}
		}
		raw = append(append(append([]byte{}, raw[:pos]...), []byte(insertion)...), raw[pos:]...)
	}

	f.Pkg.Store(path, raw)

	return nil
}

// slideLayoutPhTypeSldNum is the value of the type attribute of the p:ph
// element of the slide number placeholder of the slide layout.
const slideLayoutPhTypeSldNum = "sldNum"

// mcAlternateContentRe matches the mc:AlternateContent element (the markup
// compatibility element of the layout, for example the p14 extension of the
// layout transition): the content of the element is preserved as-is on the
// layout round trip.
var mcAlternateContentRe = regexp.MustCompile(`(?s)<mc:AlternateContent.*?</mc:AlternateContent>`)

// ensureLayoutSlideNumber provides a function to ensure the slide number
// placeholder (p:ph type="sldNum") with the slide number field (a:fld
// type="slidenum") in the slide layout, apply the options to the placeholder
// (the position and the properties of the number text) and save the modified
// layout part.
func (f *File) ensureLayoutSlideNumber(path string, options SlideNumberOptions) error {
	raw := f.readBytes(path)

	ds := new(decodeSlideLayout)
	if err := f.xmlNewDecoder(bytes.NewReader(raw)).Decode(ds); err != nil && err != io.EOF {
		return err
	}
	if err := f.layoutCheckModifiable(path, ds); err != nil {
		return err
	}

	var found *decodeShape
	for i := range ds.CommonSlideData.ShapeTree.Children {
		child := &ds.CommonSlideData.ShapeTree.Children[i]
		nv := child.Shape.NonVisualShapeProperties
		if nv == nil || nv.NonVisualProperties == nil || nv.NonVisualProperties.Ph == nil ||
			nv.NonVisualProperties.Ph.Type == nil || *nv.NonVisualProperties.Ph.Type != slideLayoutPhTypeSldNum {
			continue
		}
		found = child.Shape
		break
	}

	// the position of the placeholder: the options are applied to the
	// placeholder (the new or the existing one)
	modified := false
	if options.Xfrm != nil {
		if found == nil || found.ShapeProperties == nil {
			if found == nil {
				found = f.newSlideNumberPlaceholderShape(ds)
				ds.CommonSlideData.ShapeTree.Children = append(
					ds.CommonSlideData.ShapeTree.Children, decodeShapeTreeChild{Shape: found})
			} else {
				found.ShapeProperties = new(DecodeShapeProperties)
			}
		}
		if found.ShapeProperties.Xfrm != options.Xfrm {
			found.ShapeProperties.Xfrm = options.Xfrm
			modified = true
		}
	}

	// the slide number field of the placeholder
	if found == nil {
		found = f.newSlideNumberPlaceholderShape(ds, options.Run)
		ds.CommonSlideData.ShapeTree.Children = append(
			ds.CommonSlideData.ShapeTree.Children, decodeShapeTreeChild{Shape: found})
		modified = true
	} else {
		modified = f.ensureSlideNumberField(found, options.Run)
	}

	if !modified {
		// the placeholder with the field is already in the layout and the
		// options are empty: the layout is not modified
		return nil
	}

	return f.saveSlideLayout(path, ds)
}

// newSlideNumberPlaceholderShape provides a function to create the shape of
// the slide number placeholder of the slide layout (p:ph type="sldNum")
// with the slide number field (a:fld type="slidenum"). The id of the shape
// is the next free id of the layout, the index of the placeholder is the
// standard 12 (or the next free index when 12 is used by the other
// placeholder).
func (f *File) newSlideNumberPlaceholderShape(ds *decodeSlideLayout, run ...*DecodeRunProperties) *decodeShape {
	var runProps *DecodeRunProperties
	if len(run) > 0 && run[0] != nil {
		runProps = run[0]
	}

	idx := layoutSlideNumberIndex(ds)
	field := NewSlideNumberField()
	if runProps != nil {
		field.RunProperties = runProps
	}

	phType := slideLayoutPhTypeSldNum
	return &decodeShape{
		NonVisualShapeProperties: &decodeNonVisualShapeProperties{
			CommonNonVisualProperties: &decodeCommonNonVisualProperties{
				ID:   layoutNextShapeID(ds),
				Name: "Slide Number Placeholder " + strconv.Itoa(idx),
			},
			CommonNonVisualShapeProperties: new(decodeCommonNonVisualShapeProperties),
			NonVisualProperties:            &decodeNonVisualProperties{Ph: &Ph{Idx: &idx, Type: &phType}},
		},
		ShapeProperties: new(DecodeShapeProperties),
		TextBody: &DecodeTextBody{
			BodyProperties: &DecodeBodyProperties{},
			Paragraph:      []DecodeParagraph{{Field: []DecodeField{field}}},
		},
	}
}

// ensureSlideNumberField provides a function to ensure the slide number
// field (a:fld type="slidenum") in the text body of the slide number
// placeholder, with the given run properties of the field. It returns true
// when the text body is modified.
func (f *File) ensureSlideNumberField(shape *decodeShape, run *DecodeRunProperties) bool {
	tb := shape.TextBody
	if tb == nil {
		tb = &DecodeTextBody{
			BodyProperties: &DecodeBodyProperties{},
			Paragraph:      []DecodeParagraph{{}},
		}
		shape.TextBody = tb
	}

	var field *DecodeField
	for i := range tb.Paragraph {
		for j := range tb.Paragraph[i].Field {
			if tb.Paragraph[i].Field[j].Type == FieldSlideNumber {
				field = &tb.Paragraph[i].Field[j]
			}
		}
	}
	if field != nil {
		if run == nil {
			return false
		}
		field.RunProperties = run
		return true
	}

	// the placeholder has no the slide number field: a new paragraph with
	// the field is added (the properties of the first paragraph are kept)
	fld := NewSlideNumberField()
	if run != nil {
		fld.RunProperties = run
	}
	if len(tb.Paragraph) == 0 {
		tb.Paragraph = []DecodeParagraph{{}}
	}
	tb.Paragraph[0] = DecodeParagraph{
		ParagraphProperties: tb.Paragraph[0].ParagraphProperties,
		Field:               []DecodeField{fld},
	}

	return true
}

// layoutNextShapeID provides a function to get the next free id of the
// shape of the slide layout.
func layoutNextShapeID(ds *decodeSlideLayout) int {
	maxID := 0
	for i := range ds.CommonSlideData.ShapeTree.Children {
		child := &ds.CommonSlideData.ShapeTree.Children[i]
		nv := child.Shape.NonVisualShapeProperties
		if nv == nil || nv.CommonNonVisualProperties == nil {
			continue
		}
		if id := nv.CommonNonVisualProperties.ID; id > maxID {
			maxID = id
		}
	}

	return maxID + 1
}

// layoutSlideNumberIndex provides a function to get the index of the slide
// number placeholder of the slide layout: the standard index 12 when it is
// free, or the next free index after the indexes of the other placeholders.
func layoutSlideNumberIndex(ds *decodeSlideLayout) int {
	used := map[int]bool{}
	for i := range ds.CommonSlideData.ShapeTree.Children {
		child := &ds.CommonSlideData.ShapeTree.Children[i]
		nv := child.Shape.NonVisualShapeProperties
		if nv == nil || nv.NonVisualProperties == nil || nv.NonVisualProperties.Ph == nil ||
			nv.NonVisualProperties.Ph.Idx == nil {
			continue
		}
		used[*nv.NonVisualProperties.Ph.Idx] = true
	}
	const standardIdx = 12
	if !used[standardIdx] {
		return standardIdx
	}
	idx := 1
	for used[idx] {
		idx++
	}

	return idx
}

// slideSldNumPhRe matches the p:ph element of the slide number placeholder
// of the slide (the type attribute of the p:ph element).
var slideSldNumPhRe = regexp.MustCompile(`<p:ph[^>]*type="sldNum"`)

// slideLayoutPhElementRe matches the p:ph elements of the shape tree.
var slideLayoutPhElementRe = regexp.MustCompile(`<p:ph[^>]*/?>`)

// slideLayoutPhIdxRe matches the idx attribute of the p:ph element.
var slideLayoutPhIdxRe = regexp.MustCompile(`idx="(\d+)"`)

// slideNumbersExplicitlyEnabled reports whether the slide numbers are
// explicitly enabled by the sldNum flag of the p:hf element of the first
// slide master (the flag of the value "1" or "true"; the absent flag is the
// schema default, it does not count as enabled).
func (f *File) slideNumbersExplicitlyEnabled() bool {
	paths := f.slideMasterPaths()
	if len(paths) == 0 {
		return false
	}
	tag := masterHFElementRe.Find(f.readBytes(paths[0]))
	if tag == nil {
		return false
	}
	m := masterHFSldNumRe.FindSubmatch(tag)
	if m == nil {
		return false
	}
	v := string(m[1])

	return v == "1" || v == "true"
}

// ensureSlideSlideNumber provides a function to ensure the slide number
// placeholder (p:ph type="sldNum") of the slide: the placeholder with the
// slide number field (a:fld type="slidenum") of the same id and the same
// cached value as the field of the slide number placeholder of the layout of
// the slide (the same as the files of Google Slides). The position and the
// format of the number are inherited from the placeholder of the layout of
// the slide. The part of the slide is not modified when the placeholder is
// already in the slide.
func (f *File) ensureSlideSlideNumber(slideID int) error {
	path, ok := f.getSlideXMLPath(slideID)
	if !ok {
		return ErrSlideNotExist{slideID}
	}

	// the slide already has the placeholder: the part is not modified
	if slideSldNumPhRe.Match(f.readBytes(path)) {
		return nil
	}

	slide, err := f.slideReader(slideID)
	if err != nil {
		return err
	}

	idx := f.slideLayoutSlideNumberIndex(path)

	// the field of the placeholder of the slide: the id and the cached value
	// of the field of the placeholder of the layout of the slide (the field
	// of the slide is the same field as the field of the layout, the same as
	// the files of Google Slides); a new field when the field of the layout
	// is not found
	field := NewSlideNumberField()
	field.Text = "‹#›"
	if block, ok := f.slideLayoutSldNumPhBlock(path); ok {
		if m := slideLayoutSldNumFieldIDRe.FindStringSubmatch(block); m != nil {
			field.ID = m[1]
		}
		if m := slideLayoutSldNumFieldTextRe.FindStringSubmatch(block); m != nil {
			if text := html.UnescapeString(m[1]); text != "" {
				field.Text = text
			}
		}
	}

	phType := slideLayoutPhTypeSldNum
	slide.CommonSlideData.ShapeTree.Children = append(
		slide.CommonSlideData.ShapeTree.Children,
		decodeShapeTreeChild{Shape: &decodeShape{
			NonVisualShapeProperties: &decodeNonVisualShapeProperties{
				CommonNonVisualProperties: &decodeCommonNonVisualProperties{
					ID: f.nextShapeID(slide),
				},
				CommonNonVisualShapeProperties: new(decodeCommonNonVisualShapeProperties),
				NonVisualProperties:            &decodeNonVisualProperties{Ph: &Ph{Idx: &idx, Type: &phType}},
			},
			ShapeProperties: new(DecodeShapeProperties),
			TextBody: &DecodeTextBody{
				BodyProperties: &DecodeBodyProperties{},
				Paragraph:      []DecodeParagraph{{Field: []DecodeField{field}}},
			},
		}},
	)

	return nil
}

// slideLayoutSldNumFieldIDRe matches the id attribute of the slide number
// field (the a:fld type="slidenum" element).
var slideLayoutSldNumFieldIDRe = regexp.MustCompile(`<a:fld id="([^"]*)" type="slidenum">`)

// slideLayoutSldNumFieldTextRe matches the cached value (the a:t element) of
// the slide number field (the a:fld type="slidenum" element).
var slideLayoutSldNumFieldTextRe = regexp.MustCompile(`<a:fld id="[^"]*" type="slidenum">.*?<a:t>([^<]*)</a:t>`)

// slideLayoutPartPath provides the part path of the layout of the slide (the
// slideLayout relationship of the slide rels). It returns the empty string
// when the layout relationship of the slide is not found.
func (f *File) slideLayoutPartPath(slidePath string) string {
	rels, err := f.relsReader(f.getSlideRelsPath(slidePath))
	if err != nil || rels == nil {
		return ""
	}
	for i := range rels.Relationships {
		rel := &rels.Relationships[i]
		if rel.Type == SourceRelationshipSlideLayout {
			return "ppt/slideLayouts/" + filepath.Base(rel.Target)
		}
	}

	return ""
}

// slideLayoutSldNumPhBlock provides the block of the slide number
// placeholder (the p:sp element with the p:ph type="sldNum" element) of the
// layout of the slide. It returns ok=false when the layout of the slide or
// the placeholder of the layout is not found.
func (f *File) slideLayoutSldNumPhBlock(slidePath string) (block string, ok bool) {
	layoutPath := f.slideLayoutPartPath(slidePath)
	if layoutPath == "" {
		return "", false
	}
	s := string(f.readBytes(layoutPath))
	i := strings.Index(s, `type="sldNum"`)
	if i == -1 {
		return "", false
	}
	start := strings.LastIndex(s[:i], "<p:sp>")
	end := strings.Index(s[i:], "</p:sp>")
	if start == -1 || end == -1 {
		return "", false
	}

	return s[start : i+end+7], true
}

// slideLayoutSlideNumberIndex provides a function to get the index of the
// slide number placeholder of the layout of the slide (the placeholder of
// the slide gets the same index, so the placeholder of the slide inherits
// the properties of the placeholder of the layout). It returns the standard
// index 12 when the layout of the slide or the placeholder of the layout is
// not found.
func (f *File) slideLayoutSlideNumberIndex(slidePath string) int {
	block, ok := f.slideLayoutSldNumPhBlock(slidePath)
	if !ok {
		return 12
	}
	for _, tag := range slideLayoutPhElementRe.FindAllString(block, -1) {
		if !strings.Contains(tag, `type="sldNum"`) {
			continue
		}
		if m := slideLayoutPhIdxRe.FindStringSubmatch(tag); m != nil {
			if idx, err := strconv.Atoi(m[1]); err == nil {
				return idx
			}
		}
	}

	return 12
}

// saveSlideLayout provides a function to save the slide layout part after
// the serialize structure (the same way the slideWriter saves the slide
// part). The root element of the part keeps its origin attributes (the
// namespaces and the attributes, for example the "type" attribute of the
// layout).
func (f *File) saveSlideLayout(path string, ds *decodeSlideLayout) error {
	if _, ok := f.xmlAttr.Load(path); !ok {
		f.xmlAttr.Store(path, getRootElement(f.xmlNewDecoder(bytes.NewReader(f.readBytes(path)))))
	}

	var altContent *alternateContent
	if ds.DecodeAlternateContent != nil {
		altContent = &alternateContent{
			Content: ds.DecodeAlternateContent.Content,
			XMLNSMC: SourceRelationshipCompatibility.Value,
		}
	}

	var clrMapOverride *ColorMapOverride
	if ds.ColorMapOverride != nil {
		clrMapOverride = &ColorMapOverride{MasterColorMapping: ds.ColorMapOverride.MasterClrMapping}
	}

	var (
		shapes []Shape
		order  []shapeTreeOrderEntry
	)
	for i := range ds.CommonSlideData.ShapeTree.Children {
		child := &ds.CommonSlideData.ShapeTree.Children[i]
		if child.Shape != nil {
			shapes = append(shapes, newShape(*child.Shape))
			order = append(order, shapeTreeOrderEntry{kind: shapeTreeChildSp, index: len(shapes) - 1})
		}
	}

	output, err := xml.Marshal(&slideLayoutPart{
		XMLNSA:   NameSpaceDrawingML.Value,
		XMLNSP:   NameSpacePresentationML.Value,
		XMLNSR:   SourceRelationship.Value,
		XMLNSP14: NameSpacePowerPointR14.Value,
		XMLNSP15: NameSpacePowerPointR15.Value,
		XMLNSMC:  SourceRelationshipCompatibility.Value,
		CommonSlideData: SlideData{
			Name:       ds.CommonSlideData.Name,
			Background: newBG(ds.CommonSlideData.Background),
			ShapeTree: ShapeTree{
				NonVisualGroupShapeProperties: newNonVisualGroupShapeProperties(ds.CommonSlideData.ShapeTree.NonVisualGroupShapeProperties),
				GroupShapeProperties:          newGroupShapeProperties(ds.CommonSlideData.ShapeTree.GroupShapeProperties),
				Shape:                         shapes,
				shapeOrder:                    order,
			},
			ExtLst: newSlideDataExtLst(ds.CommonSlideData.ExtLst),
		},
		ColorMapOverride: clrMapOverride,
		AlternateContent: altContent,
		ExtLst:           ds.ExtLst,
	})
	if err != nil {
		return err
	}

	f.saveFileList(path, f.replaceNameSpaceBytes(path, output))

	return nil
}
