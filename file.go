// Copyright 2026 kenny-not-dead. All rights reserved.
// Use of this source code is governed by a BSD-style license
// that can be found in the LICENSE file.
//
// Package gopptx provides functionality to create and manipulate PowerPoint
// (.pptx) files in Go, using the Office Open XML (ECMA-376) format.

package gopptx

import (
	"bytes"
	_ "embed"
	"encoding/xml"
	"fmt"
	"io"
	"math"
	"os"
	"path/filepath"
	"sort"
	"strings"
	"sync"
)

// NewFile provides a function to create new file by default template.
// For example:
//
//	f := NewFile()
func NewFile(opts ...Options) *File {
	f := newFile()
	f.Pkg.Store(defaultXMLPathRels, []byte(xml.Header+templateRels))
	f.Pkg.Store(defaultXMLPathDocPropsApp, []byte(xml.Header+templateDocpropsApp))
	f.Pkg.Store(defaultXMLPathDocPropsCore, []byte(xml.Header+templateDocpropsCore))
	f.Pkg.Store(defaultXMLPathPresentationRels, []byte(xml.Header+templatePresentationRels))
	f.Pkg.Store(defaultXMLPathTheme, []byte(xml.Header+templateTheme))
	f.Pkg.Store(defaultXMLPathSlideMaster, []byte(xml.Header+templateSlideMaster))
	f.Pkg.Store(defaultXMLPathSlideMasterRels, []byte(xml.Header+templateSlideMasterRels))
	f.Pkg.Store(defaultXMLPathSlideLayout, []byte(xml.Header+templateSlideLayout))
	f.Pkg.Store(defaultXMLPathSlideLayoutRels, []byte(xml.Header+templateSlideLayoutRels))
	f.Pkg.Store(defaultXMLPathSlide, []byte(xml.Header+TemplateSlide))
	f.Pkg.Store(defaultXMLPathSlideRels, []byte(xml.Header+TemplateSlideRels))
	f.Pkg.Store(defaultXMLPathPresentation, []byte(xml.Header+templatePresentation))
	f.Pkg.Store(defaultXMLPathPresProps, []byte(xml.Header+templatePresProps))
	f.Pkg.Store(defaultXMLPathContentTypes, []byte(xml.Header+templateContentTypes))
	f.SlideCount = 1
	f.ContentTypes, _ = f.contentTypesReader()
	f.Presentation, _ = f.presentationReader()
	f.Relationships = sync.Map{}
	rels, _ := f.relsReader(defaultXMLPathPresentationRels)
	f.Relationships.Store(defaultXMLPathPresentationRels, rels)
	f.slideMap[defaultXMLSlideID] = defaultXMLPathSlide
	slide, _ := f.slideReader(defaultXMLSlideID)
	f.Slide.Store(defaultXMLPathSlide, slide)
	f.Theme, _ = f.themeReader()
	f.options = f.getOptions(opts...)
	return f
}

// Save provides a function to override the presentation with origin path.
func (f *File) Save(opts ...Options) error {
	if f.Path == "" {
		return ErrSave
	}
	for i := range opts {
		f.options = &opts[i]
	}
	return f.SaveAs(f.Path, *f.options)
}

// SaveAs provides a function to create or update to a presentation at the provided path.
func (f *File) SaveAs(name string, opts ...Options) error {
	if len(name) > MaxFilePathLength {
		return ErrMaxFilePathLength
	}
	f.Path = name
	if _, ok := supportedContentTypes[strings.ToLower(filepath.Ext(f.Path))]; !ok {
		return ErrPresentationFileFormat
	}
	file, err := os.OpenFile(filepath.Clean(name), os.O_WRONLY|os.O_TRUNC|os.O_CREATE, os.ModePerm)
	if err != nil {
		return err
	}
	defer file.Close()
	return f.Write(file, opts...)
}

// Close closes and cleanup the open temporary file for the presentation.
func (f *File) Close() error {
	var firstErr error

	for _, stream := range f.streams {
		_ = stream.rawData.Close()
	}
	f.streams = nil
	f.tempFiles.Range(func(k, v interface{}) bool {
		if err := os.Remove(v.(string)); err != nil && firstErr == nil {
			firstErr = err
		}
		return true
	})
	f.tempFiles.Clear()

	return firstErr
}

// Write provides a function to write to an io.Writer.
func (f *File) Write(w io.Writer, opts ...Options) error {
	_, err := f.WriteTo(w, opts...)
	return err
}

// WriteTo implements io.WriterTo to write the file.
func (f *File) WriteTo(w io.Writer, opts ...Options) (int64, error) {
	for i := range opts {
		f.options = &opts[i]
	}
	if len(f.Path) != 0 {
		contentType, ok := supportedContentTypes[strings.ToLower(filepath.Ext(f.Path))]
		if !ok {
			return 0, ErrPresentationFileFormat
		}
		if err := f.setContentTypePartProjectExtensions(contentType); err != nil {
			return 0, err
		}
	}
	buf, err := f.WriteToBuffer()
	if err != nil {
		return 0, err
	}
	return buf.WriteTo(w)
}

// WriteToBuffer provides a function to get bytes.Buffer from the saved file,
// and it allocates space in memory. Be careful when the file size is large.
func (f *File) WriteToBuffer() (*bytes.Buffer, error) {
	buf := new(bytes.Buffer)
	zw := f.ZipWriter(buf)

	if err := f.writeToZip(zw); err != nil {
		_ = zw.Close()
		return buf, err
	}
	if err := zw.Close(); err != nil {
		return buf, err
	}

	return buf, nil
}

// writeToZip provides a function to write to ZipWriter.
func (f *File) writeToZip(zw ZipWriter) error {
	f.contentTypesWriter()
	f.presentationWriter()
	// TODO: MasterWritter
	// TODO: MasterLayoutWritter
	f.slideWriter()
	f.notesSlideWriter()
	f.relsWriter()
	f.themeWriter()

	for path, stream := range f.streams {
		fi, err := zw.Create(path)
		if err != nil {
			return err
		}
		var from io.Reader
		if from, err = stream.rawData.Reader(); err != nil {
			_ = stream.rawData.Close()
			return err
		}
		written, err := io.Copy(fi, from)
		if err != nil {
			return err
		}
		if written > math.MaxUint32 {
			f.zip64Entries = append(f.zip64Entries, path)
		}
	}
	var (
		n                int
		err              error
		files, tempFiles []string
	)
	f.Pkg.Range(func(path, content interface{}) bool {
		if _, ok := f.streams[path.(string)]; ok {
			return true
		}
		files = append(files, path.(string))
		return true
	})
	sort.Sort(sort.Reverse(sort.StringSlice(files)))
	for _, path := range files {
		var fi io.Writer
		if fi, err = zw.Create(path); err != nil {
			break
		}
		content, _ := f.Pkg.Load(path)
		if n, err = fi.Write(content.([]byte)); int64(n) > math.MaxUint32 {
			f.zip64Entries = append(f.zip64Entries, path)
		}
	}
	f.tempFiles.Range(func(path, content interface{}) bool {
		if _, ok := f.Pkg.Load(path); ok {
			return true
		}
		tempFiles = append(tempFiles, path.(string))
		return true
	})
	sort.Sort(sort.Reverse(sort.StringSlice(tempFiles)))
	for _, path := range tempFiles {
		var fi io.Writer
		if fi, err = zw.Create(path); err != nil {
			break
		}
		if n, err = fi.Write(f.readBytes(path)); int64(n) > math.MaxUint32 {
			f.zip64Entries = append(f.zip64Entries, path)
		}
	}
	return err
}

// setContentTypePartProjectExtensions provides a function to set the content
// type for relationship parts and the main document part.
func (f *File) setContentTypePartProjectExtensions(contentType string) error {
	content, err := f.contentTypesReader()
	if err != nil {
		return err
	}
	content.mu.Lock()
	defer content.mu.Unlock()

	for idx, o := range content.Overrides {
		if o.PartName == "/"+defaultXMLPathPresentation {
			content.Overrides[idx].ContentType = contentType
		}
	}

	return err
}

// contentTypesWriter provides a function to save [Content_Types].xml after
// serialize structure.
func (f *File) contentTypesWriter() {
	if f.ContentTypes == nil {
		return
	}
	// keep the origin part bytes when the content types are not changed, so
	// the saved file stays identical to the origin one
	unchanged := false
	f.ContentTypes.mu.Lock()
	if content, ok := f.Pkg.Load(defaultXMLPathContentTypes); ok {
		c := contentTypes{}
		if err := xml.Unmarshal(namespaceStrictToTransitional(content.([]byte)), &c); err == nil {
			unchanged = contentTypesEqual(f.ContentTypes, &c)
		}
	}
	f.ContentTypes.mu.Unlock()
	if unchanged {
		return
	}
	output, _ := xml.Marshal(f.ContentTypes)
	f.saveFileList(defaultXMLPathContentTypes, output)
}

// contentTypesEqual reports whether the two content types structures are
// equal.
func contentTypesEqual(a, b *contentTypes) bool {
	if len(a.Defaults) != len(b.Defaults) || len(a.Overrides) != len(b.Overrides) {
		return false
	}
	for i := range a.Defaults {
		if a.Defaults[i] != b.Defaults[i] {
			return false
		}
	}
	for i := range a.Overrides {
		if a.Overrides[i] != b.Overrides[i] {
			return false
		}
	}
	return true
}

// presentationWriter provides a function to save presentation.xml after serialize structure.
func (f *File) presentationWriter() {
	newSlideID := func(s decodeSlideID) slideID {
		return slideID{
			SlideID:        s.SlideID,
			RelationshipID: s.RelationshipID,
		}
	}

	if f.Presentation != nil {
		if f.Presentation.DecodeAlternateContent != nil {
			f.Presentation.AlternateContent = &alternateContent{
				Content: f.Presentation.DecodeAlternateContent.Content,
				XMLNSMC: SourceRelationshipCompatibility.Value,
			}
		}
		f.Presentation.DecodeAlternateContent = nil

		slides := make([]slideID, len(f.Presentation.Slides.Slide))

		for i, s := range f.Presentation.Slides.Slide {
			slides[i] = newSlideID(s)
		}

		var notesMaster *notesMasterList
		if len(f.Presentation.NotesMaster.NotesMaster.RelationshipID) > 0 {
			notesMaster = &notesMasterList{
				NotesMaster: notesMasterID{
					RelationshipID: f.Presentation.NotesMaster.NotesMaster.RelationshipID,
				},
			}
		}

		output, _ := xml.Marshal(&presentation{
			XMLName:  f.Presentation.XMLName,
			XMLNSA:   NameSpaceDrawingML.Value,
			XMLNSP:   NameSpacePresentationML.Value,
			XMLNSR:   SourceRelationship.Value,
			XMLNSP14: NameSpacePowerPointR14.Value,
			XMLNSP15: NameSpacePowerPointR15.Value,
			XMLNSMC:  SourceRelationshipCompatibility.Value,
			MasterSlide: masterSlideList{
				MasterSlide: newSlideID(f.Presentation.MasterSlide.MasterSlide),
			},
			NotesMaster:      notesMaster,
			Slides:           &slideList{Slide: slides},
			SlideSize:        f.Presentation.SlideSize,
			NotesSize:        f.Presentation.NotesSize,
			DefaultTextStyle: f.Presentation.DefaultTextStyle,
			ExtList:          f.Presentation.ExtList,
		})
		f.saveFileList(f.getPresentationPath(), replaceRelationshipsBytes(f.replaceNameSpaceBytes(f.getPresentationPath(), output)))
	}
}

// replaceRelationshipsBytes; Some tools that read presentation files have very
// strict requirements about the structure of the input XML. This function is
// a horrible hack to fix that after the XML marshalling is completed.
func replaceRelationshipsBytes(content []byte) []byte {
	sourceXmlns := []byte(`xmlns:relationships="http://schemas.openxmlformats.org/officeDocument/2006/relationships" relationships`)
	targetXmlns := []byte("r")
	return bytesReplace(content, sourceXmlns, targetXmlns, -1)
}

// replaceNameSpaceBytes provides a function to replace the XML root element
// attributes by the given component part path and XML content. The original
// root element attributes (namespaces and other attributes) captured while
// reading the part are restored as is, so the serialized part keeps its
// origin root element. A namespace is added to the restored root element by
// the addNameSpaces function only when a library feature needs it.
func (f *File) replaceNameSpaceBytes(path string, content []byte) []byte {
	attrs, ok := f.xmlAttr.Load(path)
	if !ok {
		return content
	}
	attrList := attrs.([]xml.Attr)
	if len(attrList) == 0 {
		return content
	}
	idx := bytes.IndexByte(content, '>')
	if idx == -1 {
		return content
	}
	nameEnd := bytes.IndexAny(content[1:idx], " /")
	if nameEnd == -1 {
		return content
	}
	result := make([]byte, 0, len(content))
	result = append(result, content[:nameEnd+1]...)
	result = append(result, ' ')
	result = append(result, genXMLNamespace(attrList)...)
	result = append(result, content[idx+1:]...)
	return result
}

// genXMLNamespace generate serialized XML attributes with a multi namespace
// by given element attributes.
func genXMLNamespace(attr []xml.Attr) string {
	var rootElement string
	for _, v := range attr {
		if lastSpace := getXMLNamespace(v.Name.Space, attr); lastSpace != "" {
			if lastSpace == NameSpaceXML {
				lastSpace = "xml"
			}
			rootElement += fmt.Sprintf("%s:%s=\"%s\" ", lastSpace, v.Name.Local, v.Value)
			continue
		}
		rootElement += fmt.Sprintf("%s=\"%s\" ", v.Name.Local, v.Value)
	}
	return strings.TrimSpace(rootElement) + ">"
}

// themeWriter provides a function to save xl/theme/theme1.xml after serialize
// structure.
func (f *File) themeWriter() {
	newColor := func(c *decodeComplexTypeColor) complexTypeColor {
		return complexTypeColor{
			ScrgbColor:  c.ScrgbColor,
			SrgbColor:   c.SrgbColor,
			HslColor:    c.HslColor,
			SystemColor: c.SystemColor,
			SchemeColor: c.SchemeColor,
			PresetColor: c.PresetColor,
		}
	}
	newFontScheme := func(c *decodeFontCollection) fontCollection {
		return fontCollection{
			Latin: c.Latin,
			Ea:    c.Ea,
			Cs:    c.Cs,
			Font:  c.Font,
		}
	}
	if f.Theme != nil {
		output, _ := xml.Marshal(theme{
			XMLNSa: NameSpaceDrawingML.Value,
			XMLNSr: SourceRelationship.Value,
			Name:   f.Theme.Name,
			ThemeElements: baseStyles{
				ColorScheme: colorScheme{
					Name:     f.Theme.ThemeElements.ColorScheme.Name,
					Dk1:      newColor(&f.Theme.ThemeElements.ColorScheme.Dk1),
					Lt1:      newColor(&f.Theme.ThemeElements.ColorScheme.Lt1),
					Dk2:      newColor(&f.Theme.ThemeElements.ColorScheme.Dk2),
					Lt2:      newColor(&f.Theme.ThemeElements.ColorScheme.Lt2),
					Accent1:  newColor(&f.Theme.ThemeElements.ColorScheme.Accent1),
					Accent2:  newColor(&f.Theme.ThemeElements.ColorScheme.Accent2),
					Accent3:  newColor(&f.Theme.ThemeElements.ColorScheme.Accent3),
					Accent4:  newColor(&f.Theme.ThemeElements.ColorScheme.Accent4),
					Accent5:  newColor(&f.Theme.ThemeElements.ColorScheme.Accent5),
					Accent6:  newColor(&f.Theme.ThemeElements.ColorScheme.Accent6),
					Hlink:    newColor(&f.Theme.ThemeElements.ColorScheme.Hlink),
					FolHlink: newColor(&f.Theme.ThemeElements.ColorScheme.FolHlink),
				},
				FontScheme: fontScheme{
					Name:      f.Theme.ThemeElements.FontScheme.Name,
					MajorFont: newFontScheme(&f.Theme.ThemeElements.FontScheme.MajorFont),
					MinorFont: newFontScheme(&f.Theme.ThemeElements.FontScheme.MinorFont),
				},
				FormatScheme: styleMatrix{
					Name:            f.Theme.ThemeElements.FormatScheme.Name,
					FillStyleList:   f.Theme.ThemeElements.FormatScheme.FillStyleList,
					LineStyleList:   f.Theme.ThemeElements.FormatScheme.LineStyleList,
					EffectStyleList: f.Theme.ThemeElements.FormatScheme.EffectStyleList,
					BgFillStyleList: f.Theme.ThemeElements.FormatScheme.BgFillStyleList,
				},
			},
			ObjectDefaults:    f.Theme.ObjectDefaults,
			ExtraClrSchemeLst: f.Theme.ExtraClrSchemeLst,
		})
		f.saveFileList(defaultXMLPathTheme, f.replaceNameSpaceBytes(defaultXMLPathTheme, output))
	}
}

// relsWriter provides a function to save relationships after serialize structure.
func (f *File) relsWriter() {
	f.Relationships.Range(func(path, rel interface{}) bool {
		if rel != nil {
			rels := rel.(*relationships)
			// keep the origin part bytes when the relationship list is not
			// changed, so the saved file stays identical to the origin one
			rels.mu.Lock()
			unchanged := false
			if content, ok := f.Pkg.Load(path.(string)); ok {
				c := relationships{}
				if err := xml.Unmarshal(namespaceStrictToTransitional(content.([]byte)), &c); err == nil {
					unchanged = relationshipsEqual(rels.Relationships, c.Relationships)
				}
			}
			rels.mu.Unlock()
			if unchanged {
				return true
			}
			output, _ := xml.Marshal(rels)
			if strings.HasPrefix(path.(string), "ppt/slides/_rels/slide") {
				output = f.replaceNameSpaceBytes(path.(string), output)
			}
			f.saveFileList(path.(string), replaceRelationshipsBytes(output))
		}
		return true
	})
}

// relationshipsEqual reports whether the two relationship lists are equal.
func relationshipsEqual(a, b []relationship) bool {
	if len(a) != len(b) {
		return false
	}
	for i := range a {
		if a[i] != b[i] {
			return false
		}
	}
	return true
}

// newHyperlink provides a function to convert the decode hyperlink to the
// hyperlink.
func newHyperlink(dh *DecodeHyperlink) *Hyperlink {
	if dh == nil {
		return nil
	}

	return &Hyperlink{
		ID:      dh.ID,
		Tooltip: dh.Tooltip,
	}
}

// newCommonNonVisualProperties provides a function to convert the decode
// common non-visual properties to the common non-visual properties.
func newCommonNonVisualProperties(
	dcnvp *decodeCommonNonVisualProperties,
) *CommonNonVisualProperties {
	if dcnvp == nil {
		return nil
	}

	cnvp := &CommonNonVisualProperties{
		ID:    dcnvp.ID,
		Name:  dcnvp.Name,
		Title: dcnvp.Title,
	}

	if dcnvp.HlinkClick != nil {
		cnvp.HlinkClick = &Hyperlink{
			ID:      dcnvp.HlinkClick.ID,
			Tooltip: dcnvp.HlinkClick.Tooltip,
		}
	}

	if dcnvp.HlinkMouseOver != nil {
		cnvp.HlinkMouseOver = &Hyperlink{
			ID:      dcnvp.HlinkMouseOver.ID,
			Tooltip: dcnvp.HlinkMouseOver.Tooltip,
		}
	}

	if dcnvp.ExtLst != nil {
		cnvp.ExtLst = &ExtLst{Content: dcnvp.ExtLst.Content}
	}

	return cnvp
}

// newNonVisualGroupShapeProperties provides a function to convert the decode
// non-visual group shape properties to the non-visual group shape properties.
func newNonVisualGroupShapeProperties(
	nvgsp *decodeNonVisualGroupShapeProperties,
) *NonVisualGroupShapeProperties {
	if nvgsp == nil {
		return nil
	}

	return &NonVisualGroupShapeProperties{
		CommonNonVisualProperties:           newCommonNonVisualProperties(nvgsp.CommonNonVisualProperties),
		CommonNonVisualGroupShapeProperties: nvgsp.CommonNonVisualGroupShapeProperties,
		NonVisualProperties:                 (*NonVisualProperties)(nvgsp.NonVisualProperties),
	}
}

// newXfrm provides a function to convert the decode transform to the
// transform.
func newXfrm(dx *DecodeXfrm) *Xfrm {
	if dx == nil {
		return nil
	}

	return &Xfrm{
		Offset:       dx.Offset,
		Extents:      dx.Extents,
		ChildOffset:  dx.ChildOffset,
		ChildExtents: dx.ChildExtents,
	}
}

// newGroupShapeProperties provides a function to convert the decode group
// shape properties to the group shape properties.
func newGroupShapeProperties(c *decodeGroupShapeProperties) *GroupShapeProperties {
	if c == nil {
		return nil
	}

	return &GroupShapeProperties{
		Xfrm: newXfrm(c.Xfrm),
	}
}

// newPresetGeometry provides a function to convert the decode preset geometry
// to the preset geometry.
func newPresetGeometry(dpg *DecodePresetGeometry) *PresetGeometry {
	if dpg == nil {
		return nil
	}

	return &PresetGeometry{
		Preset:          dpg.Preset,
		AdjustValueList: dpg.AdjustValueList,
	}
}

// newSolidFill provides a function to convert the decode solid fill to the
// solid fill.
func newSolidFill(dsf *DecodeSolidFill) *SolidFill {
	if dsf == nil {
		return nil
	}

	return &SolidFill{
		SolidRGBColor: dsf.SolidRGBColor,
	}
}

// newLine provides a function to convert the decode line to the line.
func newLine(dl *DecodeLine) *Line {
	if dl == nil {
		return nil
	}

	return &Line{
		Width:     dl.Width,
		NoFill:    dl.NoFill,
		SolidFill: newSolidFill(dl.SolidFill),
		PrstDash:  dl.PrstDash,
	}
}

// newRunProperties provides a function to convert the decode run properties
// to the run properties.
func newRunProperties(dr *DecodeRunProperties) *RunProperties {
	if dr == nil {
		return nil
	}

	rp := &RunProperties{
		Bold:           dr.Bold,
		Lang:           dr.Lang,
		Size:           dr.Size,
		Space:          dr.Space,
		Strike:         dr.Strike,
		Latin:          dr.Latin,
		EastAsian:      dr.EastAsian,
		ComplexScript:  dr.ComplexScript,
		HlinkClick:     newHyperlink(dr.HlinkClick),
		HlinkMouseOver: newHyperlink(dr.HlinkMouseOver),
	}

	if dr.SolidFill != nil {
		rp.SolidFill = &SolidFill{
			SolidRGBColor: dr.SolidFill.SolidRGBColor,
		}
	}

	return rp
}

// newRuns provides a function to convert the decode runs to the runs.
func newRuns(r []DecodeRuns) []Runs {
	runs := make([]Runs, len(r))
	for i, run := range r {
		runs[i] = Runs{
			RunProperties: newRunProperties(run.RunProperties),
			Text:          run.Text,
		}
	}

	return runs
}

// newParagraphProperties provides a function to convert the decode paragraph
// properties to the paragraph properties.
func newParagraphProperties(dpp *DecodeParagraphProperties) *ParagraphProperties {
	if dpp == nil {
		return nil
	}

	var lineSpacing *LineSpacing
	if dpp.LineSpacing != nil {
		lineSpacing = &LineSpacing{
			SpacingPercent: dpp.LineSpacing.SpacingPercent,
		}
	}

	newSpacingChange := func(dsc *DecodeSpacingChange) *SpacingChange {
		if dsc == nil {
			return nil
		}

		return &SpacingChange{
			SpacingPercent: dsc.SpacingPercent,
			SpacingPts:     dsc.SpacingPts,
		}
	}

	newBulletFont := func(dbf *DecodeBulletFont) *BulletFont {
		if dbf == nil {
			return nil
		}

		return &BulletFont{
			Typeface: dbf.Typeface,
		}
	}
	newBulletChar := func(dbc *DecodeBulletChar) *BulletChar {
		if dbc == nil {
			return nil
		}

		return &BulletChar{
			Char: dbc.Char,
		}
	}

	return &ParagraphProperties{
		Indent:           dpp.Indent,
		Lvl:              dpp.Lvl,
		MarL:             dpp.MarL,
		RTL:              dpp.RTL,
		Align:            dpp.Align,
		LineSpacing:      lineSpacing,
		SpacingBefore:    newSpacingChange(dpp.SpacingBefore),
		SpacingAfter:     newSpacingChange(dpp.SpacingAfter),
		BuFont:           newBulletFont(dpp.BuFont),
		BuChar:           newBulletChar(dpp.BuChar),
		BuNone:           dpp.BuNone,
		DefRunProperties: newRunProperties(dpp.DefRunProperties),
	}
}

// newTextBody provides a function to convert the decode text body to the text
// body.
func newTextBody(dt *DecodeTextBody) *TextBody {
	if dt == nil {
		return nil
	}

	paragraphs := make([]Paragraph, len(dt.Paragraph))
	for i, p := range dt.Paragraph {
		paragraphs[i] = Paragraph{
			ParagraphProperties:       newParagraphProperties(p.ParagraphProperties),
			Runs:                      newRuns(p.Runs),
			EndParagraphRunProperties: newRunProperties(p.EndParagraphRunProperties),
		}
	}

	var bodyProperties *BodyProperties
	if dt.BodyProperties != nil {
		bodyProperties = &BodyProperties{
			AnchorCtr:        dt.BodyProperties.AnchorCtr,
			Anchor:           dt.BodyProperties.Anchor,
			BIns:             dt.BodyProperties.BIns,
			LIns:             dt.BodyProperties.LIns,
			SpcFirstLastPara: dt.BodyProperties.SpcFirstLastPara,
			RIns:             dt.BodyProperties.RIns,
			Wrap:             dt.BodyProperties.Wrap,
			TIns:             dt.BodyProperties.TIns,
			NormAutofit:      dt.BodyProperties.NormAutofit,
			NoAutofit:        dt.BodyProperties.NoAutofit,
		}
	}

	return &TextBody{
		BodyProperties: bodyProperties,
		ListStyle:      dt.ListStyle,
		Paragraph:      paragraphs,
	}
}

// newShapeProperties provides a function to convert the decode shape
// properties to the shape properties.
func newShapeProperties(dsp *DecodeShapeProperties) *ShapeProperties {
	if dsp == nil {
		return nil
	}

	return &ShapeProperties{
		Xfrm:           newXfrm(dsp.Xfrm),
		CustGeom:       dsp.CustGeom,
		PresetGeometry: newPresetGeometry(dsp.PresetGeometry),
		NoFill:         dsp.NoFill,
		SolidFill:      newSolidFill(dsp.SolidFill),
		Ln:             newLine(dsp.Ln),
		EffectList:     dsp.EffectList,
	}
}

// newNonVisualShapeProperties provides a function to convert the decode
// non-visual shape properties to the non-visual shape properties.
func newNonVisualShapeProperties(dnsp *decodeNonVisualShapeProperties) *NonVisualShapeProperties {
	if dnsp == nil {
		return nil
	}

	var commonNonVisualShapeProperties *CommonNonVisualShapeProperties
	if dnsp.CommonNonVisualShapeProperties != nil {
		commonNonVisualShapeProperties = &CommonNonVisualShapeProperties{
			ShapeLocks: dnsp.CommonNonVisualShapeProperties.ShapeLocks,
			TxBox:      boolToXMLBool(dnsp.CommonNonVisualShapeProperties.TxBox),
		}
	}

	var nonVisualProperties *NonVisualProperties
	if dnsp.NonVisualProperties != nil {
		nonVisualProperties = &NonVisualProperties{
			Ph: dnsp.NonVisualProperties.Ph,
		}
	}

	return &NonVisualShapeProperties{
		CommonNonVisualProperties:      newCommonNonVisualProperties(dnsp.CommonNonVisualProperties),
		CommonNonVisualShapeProperties: commonNonVisualShapeProperties,
		NonVisualProperties:            nonVisualProperties,
	}
}

// newShape provides a function to convert the decode shape to the shape.
func newShape(ds decodeShape) Shape {
	return Shape{
		NonVisualShapeProperties: newNonVisualShapeProperties(ds.NonVisualShapeProperties),
		ShapeProperties:          newShapeProperties(ds.ShapeProperties),
		TextBody:                 newTextBody(ds.TextBody),
	}
}

// newPicture provides a function to convert the decode picture to the picture.
func newPicture(dp decodePicture) Picture {
	pic := Picture{}
	if dp.NonVisualPictureProperties != nil {
		pic.NonVisualPictureProperties = &NonVisualPictureProperties{
			CommonNonVisualProperties:        newCommonNonVisualProperties(dp.NonVisualPictureProperties.CommonNonVisualProperties),
			CommonNonVisualPictureProperties: dp.NonVisualPictureProperties.CommonNonVisualPictureProperties,
			NonVisualProperties:              (*NonVisualProperties)(dp.NonVisualPictureProperties.NonVisualProperties),
		}
	}
	if dp.BlipFill != nil {
		pic.BlipFill = &BlipFill{}
		if dp.BlipFill.Blip != nil {
			pic.BlipFill.Blip = &Blip{
				Embed:       dp.BlipFill.Blip.Embed,
				AlphaModFix: dp.BlipFill.Blip.AlphaModFix,
			}
		}
		if dp.BlipFill.Stretch != nil && dp.BlipFill.Stretch.FillRect != nil {
			pic.BlipFill.Stretch = &Stretch{
				FillRect: dp.BlipFill.Stretch.FillRect,
			}
		}
	}
	pic.ShapeProperties = newShapeProperties(dp.ShapeProperties)
	return pic
}

// newConnector provides a function to convert the decode connector to the
// connector.
func newConnector(dc decodeConnector) Connector {
	conn := Connector{}
	if dc.NonVisualConnectorProperties != nil {
		conn.NonVisualConnectorProperties = &NonVisualConnectorProperties{
			CommonNonVisualProperties:          newCommonNonVisualProperties(dc.NonVisualConnectorProperties.CommonNonVisualProperties),
			CommonNonVisualConnectorProperties: dc.NonVisualConnectorProperties.CommonNonVisualConnectorProperties,
			NonVisualProperties:                (*NonVisualProperties)(dc.NonVisualConnectorProperties.NonVisualProperties),
		}
	}
	conn.ShapeProperties = newShapeProperties(dc.ShapeProperties)

	return conn
}

// newBG provides a function to convert the decode background to the
// background.
func newBG(db *DecodeBackground) *bg {
	if db == nil {
		return nil
	}

	var bgProps *bgProperties
	if db.BGProperties != nil {
		bgProps = &bgProperties{
			NoFill:     db.BGProperties.NoFill,
			SolidFill:  newSolidFill(db.BGProperties.SolidFill),
			EffectList: db.BGProperties.EffectList,
		}
	}

	return &bg{BGProperties: bgProps}
}

// slideWriter provides a function to save ppt/slides/slide%d.xml after
// serialize structure.
func (f *File) slideWriter() {
	var (
		arr    []byte
		buffer = bytes.NewBuffer(arr)
	)
	f.Slide.Range(func(p, slide interface{}) bool {
		if slide != nil {

			ds := slide.(*decodeSlide)

			// the shapes, the pictures and the connectors of the shape tree;
			// the document order of the children (the z-order) is recorded in
			// the order of the shape tree and is preserved on save
			var (
				shapes     []Shape
				pictures   []Picture
				connectors []Connector
				order      []shapeTreeOrderEntry
			)
			for i := range ds.CommonSlideData.ShapeTree.Children {
				child := &ds.CommonSlideData.ShapeTree.Children[i]
				switch {
				case child.Shape != nil:
					shapes = append(shapes, newShape(*child.Shape))
					order = append(order, shapeTreeOrderEntry{kind: shapeTreeChildSp, index: len(shapes) - 1})
				case child.Picture != nil:
					pictures = append(pictures, newPicture(*child.Picture))
					order = append(order, shapeTreeOrderEntry{kind: shapeTreeChildPic, index: len(pictures) - 1})
				case child.Connector != nil:
					connectors = append(connectors, newConnector(*child.Connector))
					order = append(order, shapeTreeOrderEntry{kind: shapeTreeChildCxnSp, index: len(connectors) - 1})
				}
			}

			var altContent *alternateContent
			if ds.DecodeAlternateContent != nil {
				altContent = &alternateContent{
					Content: ds.DecodeAlternateContent.Content,
					XMLNSMC: SourceRelationshipCompatibility.Value,
				}
			}

			var clrMapOverride *ClrMapOverride
			if ds.ColorMapOverride != nil {
				clrMapOverride = &ClrMapOverride{
					MasterClrMapping: ds.ColorMapOverride.MasterClrMapping,
				}
			}

			output, _ := xml.Marshal(&Slide{
				XMLName:  ds.XMLName,
				XMLNSA:   NameSpaceDrawingML.Value,
				XMLNSP:   NameSpacePresentationML.Value,
				XMLNSR:   SourceRelationship.Value,
				XMLNSP14: NameSpacePowerPointR14.Value,
				XMLNSP15: NameSpacePowerPointR15.Value,
				XMLNSMC:  SourceRelationshipCompatibility.Value,
				CommonSlideData: SlideData{
					Background: newBG(ds.CommonSlideData.Background),
					ShapeTree: ShapeTree{
						NonVisualGroupShapeProperties: newNonVisualGroupShapeProperties(ds.CommonSlideData.ShapeTree.NonVisualGroupShapeProperties),
						GroupShapeProperties:          newGroupShapeProperties(ds.CommonSlideData.ShapeTree.GroupShapeProperties),
						Shape:                         shapes,
						Picture:                       pictures,
						Connector:                     connectors,
						shapeOrder:                    order,
					},
				},
				ColorMapOverride: clrMapOverride,
				AlternateContent: altContent,
			})

			f.saveFileList(p.(string), f.replaceNameSpaceBytes(p.(string), output))
			_, ok := f.checked.Load(p.(string))
			if ok {
				f.Slide.Delete(p.(string))
				f.checked.Delete(p.(string))
			}
			buffer.Reset()
		}
		return true
	})
}

// notesSlideWriter provides a function to save ppt/notesSlides/notesSlide%d.xml
// after serialize structure.
func (f *File) notesSlideWriter() {
	f.NotesSlides.Range(func(path, notes interface{}) bool {
		if notes != nil {

			ns := notes.(*decodeNotesSlide)

			var altContent *alternateContent
			if ns.DecodeAlternateContent != nil {
				altContent = &alternateContent{
					Content: ns.DecodeAlternateContent.Content,
					XMLNSMC: SourceRelationshipCompatibility.Value,
				}
			}

			var clrMapOverride *ClrMapOverride
			if ns.ColorMapOverride != nil {
				clrMapOverride = &ClrMapOverride{
					MasterClrMapping: ns.ColorMapOverride.MasterClrMapping,
				}
			}

			shapes := make([]Shape, len(ns.CommonSlideData.ShapeTree.Shape))
			for i, shape := range ns.CommonSlideData.ShapeTree.Shape {
				shapes[i] = newShape(shape)
			}

			output, _ := xml.Marshal(&NotesSlide{
				XMLName: ns.XMLName,
				XMLNSA:  NameSpaceDrawingML.Value,
				XMLNSP:  NameSpacePresentationML.Value,
				XMLNSR:  SourceRelationship.Value,
				CommonSlideData: NotesSlideData{
					ShapeTree: NotesShapeTree{
						NonVisualGroupShapeProperties: newNonVisualGroupShapeProperties(ns.CommonSlideData.ShapeTree.NonVisualGroupShapeProperties),
						GroupShapeProperties:          newGroupShapeProperties(ns.CommonSlideData.ShapeTree.GroupShapeProperties),
						Shape:                         shapes,
					},
				},
				ColorMapOverride: clrMapOverride,
				AlternateContent: altContent,
			})

			f.saveFileList(path.(string), f.replaceNameSpaceBytes(path.(string), output))
			_, ok := f.checked.Load(path.(string))
			if ok {
				f.NotesSlides.Delete(path.(string))
				f.checked.Delete(path.(string))
			}
		}
		return true
	})
}
