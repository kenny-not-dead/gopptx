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
	"io"
	"regexp"
	"strconv"
	"strings"
)

// slideLayoutRef is a slide layout of the presentation: the id of the layout
// (the id attribute of the p:sldLayoutId element of the slide master) and the
// part paths of the slide master and the slide layout.
type slideLayoutRef struct {
	id     int
	master string
	layout string
}

// sldLayoutIDTagRe matches the p:sldLayoutId element of the slide master
// (the slide layout id list).
var sldLayoutIDTagRe = regexp.MustCompile(`<[A-Za-z0-9]+:sldLayoutId\s[^>]*?>`)

// sldLayoutIDAttrRe matches the id attribute of the p:sldLayoutId element.
var sldLayoutIDAttrRe = regexp.MustCompile(` id="(\d+)"`)

// sldLayoutIDRelAttrRe matches the r:id attribute of the p:sldLayoutId element.
var sldLayoutIDRelAttrRe = regexp.MustCompile(` r:id="(rId\d+)"`)

// slideLayoutRefs provides a function to get the slide layouts of the
// presentation (the p:sldLayoutId elements of the sldLayoutIdLst element of
// each slide master), in the document order of the slide masters (the part
// order) and the layout id lists.
func (f *File) slideLayoutRefs() ([]slideLayoutRef, error) {
	var refs []slideLayoutRef
	for _, masterPath := range f.slideMasterPaths() {
		raw := f.readBytes(masterPath)
		rels, err := f.relsReader(f.getSlideRelsPath(masterPath))
		if err != nil {
			return nil, err
		}
		for _, tag := range sldLayoutIDTagRe.FindAll(raw, -1) {
			idMatch := sldLayoutIDAttrRe.FindSubmatch(tag)
			relMatch := sldLayoutIDRelAttrRe.FindSubmatch(tag)
			if idMatch == nil || relMatch == nil {
				continue
			}
			id, _ := strconv.Atoi(string(idMatch[1]))
			for _, rel := range rels.Relationships {
				if rel.ID != string(relMatch[1]) {
					continue
				}
				refs = append(refs, slideLayoutRef{
					id:     id,
					master: masterPath,
					layout: resolvePartPath(masterPath, rel.Target),
				})
			}
		}
	}

	return refs, nil
}

// layoutReader provides a function to get the parsed slide layout by the
// layout id (the same way the slideReader gets the slide by the slide id),
// and the part path of the slide layout.
func (f *File) layoutReader(layoutID int) (*decodeSlideLayout, string, error) {
	refs, err := f.slideLayoutRefs()
	if err != nil {
		return nil, "", err
	}
	for _, ref := range refs {
		if ref.id != layoutID {
			continue
		}
		ds := new(decodeSlideLayout)
		if err := f.xmlNewDecoder(bytes.NewReader(f.readBytes(ref.layout))).Decode(ds); err != nil && err != io.EOF {
			return nil, "", err
		}

		return ds, ref.layout, nil
	}

	return nil, "", ErrSlideLayoutNotExist{layoutID}
}

// layoutCheckModifiable provides a function to check the slide layout for
// the elements that are not supported by the layout modification (the
// modification of the layout is not supported for such parts, so the layout
// cannot be modified without losing the content of such elements).
func (f *File) layoutCheckModifiable(path string, ds *decodeSlideLayout) error {
	raw := f.readBytes(path)

	// the elements that are not modeled by the layout round trip (the
	// markup compatibility elements of the layout root are preserved as-is,
	// the same as in the slide round trip)
	stripped := mcAlternateContentRe.ReplaceAll(raw, nil)
	for _, element := range []string{
		":grpSp>", ":transition", ":timing", ":hf",
	} {
		if bytes.Contains(stripped, []byte(element)) {
			return ErrSlideLayoutUnsupported{
				Layout:  path,
				Element: strings.TrimPrefix(element, ":"),
			}
		}
	}
	if element := commonSlideDataUnsupportedElement(raw); element != "" {
		return ErrSlideLayoutUnsupported{Layout: path, Element: element}
	}
	for i := range ds.CommonSlideData.ShapeTree.Children {
		child := &ds.CommonSlideData.ShapeTree.Children[i]
		if child.Shape == nil {
			return ErrSlideLayoutUnsupported{Layout: path, Element: "the shape of the type other than sp"}
		}
	}

	return nil
}

// GetSlideLayoutList provides a function to get the id list of the slide
// layouts of the presentation (the id attributes of the p:sldLayoutId
// elements of the slide masters).
func (f *File) GetSlideLayoutList() (list []int) {
	refs, err := f.slideLayoutRefs()
	if err != nil {
		return nil
	}
	for _, ref := range refs {
		list = append(list, ref.id)
	}

	return list
}

// GetLayoutShapes provides a function to get the shapes of the slide layout
// by the given layout id, in the document order.
func (f *File) GetLayoutShapes(layoutID int) ([]decodeShape, error) {
	ds, _, err := f.layoutReader(layoutID)
	if err != nil {
		return nil, err
	}

	return ds.CommonSlideData.shapes(), nil
}

// CreateLayoutShape provides a function to create a new shape (p:sp) in the
// shape tree of the slide layout by the given layout id, and returns the id
// of the new shape. The new shape is appended to the end of the shape tree:
// it is drawn above the existing shapes of the layout.
func (f *File) CreateLayoutShape(layoutID int, shapeProperties DecodeShapeProperties, textBody DecodeTextBody) (int, error) {
	ds, path, err := f.layoutReader(layoutID)
	if err != nil {
		return -1, err
	}
	if err := f.layoutCheckModifiable(path, ds); err != nil {
		return -1, err
	}

	shapeID := ds.CommonSlideData.nextShapeID()
	shape := newDecodeShape(shapeID, shapeProperties, textBody)
	ds.CommonSlideData.ShapeTree.Children = append(
		ds.CommonSlideData.ShapeTree.Children, decodeShapeTreeChild{Shape: &shape})

	return shapeID, f.saveSlideLayout(path, ds)
}

// DeleteLayoutShape provides a function to delete the shape by the given
// shape id from the shape tree of the slide layout by the given layout id.
func (f *File) DeleteLayoutShape(layoutID int, shapeID int) error {
	ds, path, err := f.layoutReader(layoutID)
	if err != nil {
		return err
	}
	if err := f.layoutCheckModifiable(path, ds); err != nil {
		return err
	}
	if !ds.CommonSlideData.deleteShape(shapeID) {
		return ErrShapeNotExist{shapeID}
	}

	return f.saveSlideLayout(path, ds)
}

// SetLayoutShapeTextBody provides a function to set the shape text body by
// the given layout id and shape id, replacing the current text body of the
// shape.
func (f *File) SetLayoutShapeTextBody(layoutID int, shapeID int, textBody DecodeTextBody) error {
	ds, path, err := f.layoutReader(layoutID)
	if err != nil {
		return err
	}
	if err := f.layoutCheckModifiable(path, ds); err != nil {
		return err
	}
	if shape := ds.CommonSlideData.findShape(shapeID); shape == nil {
		return ErrShapeNotExist{shapeID}
	} else {
		shape.TextBody = &textBody
	}

	return f.saveSlideLayout(path, ds)
}

// SetLayoutShapeText provides a function to set the text of the shape of the
// slide layout by the given layout id and shape id, replacing the paragraph
// of the text body of the shape (the same way the SetShapeText function sets
// the text of the slide shape). The body properties of the current text body
// are kept.
func (f *File) SetLayoutShapeText(layoutID int, shapeID int, text string) error {
	ds, path, err := f.layoutReader(layoutID)
	if err != nil {
		return err
	}
	if err := f.layoutCheckModifiable(path, ds); err != nil {
		return err
	}
	shape := ds.CommonSlideData.findShape(shapeID)
	if shape == nil {
		return ErrShapeNotExist{shapeID}
	}
	if shape.TextBody == nil {
		shape.TextBody = &DecodeTextBody{
			BodyProperties: &DecodeBodyProperties{},
			ListStyle:      &struct{}{},
		}
	}
	shape.TextBody.Paragraph = []DecodeParagraph{newSetShapeTextParagraph(text)}

	return f.saveSlideLayout(path, ds)
}

// SetLayoutShapeXfrm provides a function to set the position and the size
// (a:xfrm) of the shape of the slide layout by the given layout id and shape id.
func (f *File) SetLayoutShapeXfrm(layoutID int, shapeID int, xfrm DecodeXfrm) error {
	ds, path, err := f.layoutReader(layoutID)
	if err != nil {
		return err
	}
	if err := f.layoutCheckModifiable(path, ds); err != nil {
		return err
	}
	shape := ds.CommonSlideData.findShape(shapeID)
	if shape == nil {
		return ErrShapeNotExist{shapeID}
	}
	if shape.ShapeProperties == nil {
		shape.ShapeProperties = &DecodeShapeProperties{}
	}
	shape.ShapeProperties.Xfrm = &xfrm

	return f.saveSlideLayout(path, ds)
}

// SetLayoutShapeFill provides a function to set the solid fill of the shape
// of the slide layout by the given layout id and shape id, replacing the
// current fill of the shape.
func (f *File) SetLayoutShapeFill(layoutID int, shapeID int, fill DecodeSolidFill) error {
	ds, path, err := f.layoutReader(layoutID)
	if err != nil {
		return err
	}
	if err := f.layoutCheckModifiable(path, ds); err != nil {
		return err
	}
	shape := ds.CommonSlideData.findShape(shapeID)
	if shape == nil {
		return ErrShapeNotExist{shapeID}
	}
	if shape.ShapeProperties == nil {
		shape.ShapeProperties = &DecodeShapeProperties{}
	}
	shape.ShapeProperties.NoFill = nil
	shape.ShapeProperties.SolidFill = &fill

	return f.saveSlideLayout(path, ds)
}

// GetLayoutShapeFill provides a function to get the solid fill of the shape
// of the slide layout by the given layout id and shape id. It returns nil
// when the shape has no solid fill.
func (f *File) GetLayoutShapeFill(layoutID int, shapeID int) (*DecodeSolidFill, error) {
	ds, _, err := f.layoutReader(layoutID)
	if err != nil {
		return nil, err
	}
	shape := ds.CommonSlideData.findShape(shapeID)
	if shape == nil {
		return nil, ErrShapeNotExist{shapeID}
	}
	if shape.ShapeProperties == nil {
		return nil, nil
	}

	return shape.ShapeProperties.SolidFill, nil
}

// SetSlideLayoutBackground provides a function to set the background of the
// slide layout by the given layout id, replacing the current background of
// the layout.
func (f *File) SetSlideLayoutBackground(layoutID int, background DecodeBackground) error {
	ds, path, err := f.layoutReader(layoutID)
	if err != nil {
		return err
	}
	if err := f.layoutCheckModifiable(path, ds); err != nil {
		return err
	}

	ds.CommonSlideData.Background = &background

	return f.saveSlideLayout(path, ds)
}

// GetSlideLayoutBackground provides a function to get the background of the
// slide layout by the given layout id. It returns nil when the layout has no
// explicit background and inherits it from the slide master.
func (f *File) GetSlideLayoutBackground(layoutID int) (*DecodeBackground, error) {
	ds, _, err := f.layoutReader(layoutID)
	if err != nil {
		return nil, err
	}

	return ds.CommonSlideData.Background, nil
}

// DeleteSlideLayoutBackground provides a function to delete the background
// of the slide layout by the given layout id, so the layout inherits the
// background from the slide master.
func (f *File) DeleteSlideLayoutBackground(layoutID int) error {
	ds, path, err := f.layoutReader(layoutID)
	if err != nil {
		return err
	}
	if err := f.layoutCheckModifiable(path, ds); err != nil {
		return err
	}

	ds.CommonSlideData.Background = nil

	return f.saveSlideLayout(path, ds)
}

// slideLayoutPartIndexRe matches the index of the part of the slide layout
// (the N of the slideLayoutN.xml part name).
var slideLayoutPartIndexRe = regexp.MustCompile(`^ppt/slideLayouts/slideLayout(\d+)\.xml$`)

// nextSlideLayoutPartIndex provides a function to get the next free index of
// the part of the slide layout (the N of the slideLayoutN.xml part name).
func (f *File) nextSlideLayoutPartIndex() int {
	next := 1
	f.Pkg.Range(func(k, _ interface{}) bool {
		name, _ := k.(string)
		if m := slideLayoutPartIndexRe.FindStringSubmatch(name); m != nil {
			if index, _ := strconv.Atoi(m[1]); index >= next {
				next = index + 1
			}
		}
		return true
	})

	return next
}

// NewSlideLayout provides a function to create a new slide layout (a copy of
// the slide layout with the given layout id) in the slide master of the
// source layout, and returns the id of the new layout. The new layout is a
// byte copy of the source layout (the shapes and the placeholders of the
// source layout, including the slide number placeholder), so the new slides
// that use the new layout inherit the design of the source layout. The part
// of the new layout, the rels part of the new layout, the content types and
// the slide layout id list of the slide master are updated automatically.
//
//	layoutID, err := f.NewSlideLayout(f.GetSlideLayoutList()[0])
func (f *File) NewSlideLayout(layoutID int) (int, error) {
	refs, err := f.slideLayoutRefs()
	if err != nil {
		return -1, err
	}
	var source *slideLayoutRef
	nextLayoutID := 0
	for i := range refs {
		// the layout ids are positive (the next id is the max id + 1)
		if refs[i].id >= nextLayoutID {
			nextLayoutID = refs[i].id + 1
		}
		if refs[i].id == layoutID {
			source = &refs[i]
		}
	}
	if source == nil {
		return -1, ErrSlideLayoutNotExist{layoutID}
	}

	index := f.nextSlideLayoutPartIndex()
	layoutPath := "ppt/slideLayouts/slideLayout" + strconv.Itoa(index) + ".xml"
	relsPath := f.getSlideRelsPath(layoutPath)

	// the part and the rels part of the new layout: the byte copy of the
	// part and the rels part of the source layout (the rels part of the
	// layout points to the slide master of the layout, the new layout is
	// added to the same slide master)
	f.Pkg.Store(layoutPath, f.readBytes(source.layout))
	if sourceRels := f.readBytes(f.getSlideRelsPath(source.layout)); len(sourceRels) > 0 {
		f.Pkg.Store(relsPath, sourceRels)
	} else {
		// the source layout has no rels part (an invalid package, the
		// layout must point to the slide master): the new layout gets the
		// empty relationships part
		relsOutput, _ := xml.Marshal(&relationships{})
		f.Pkg.Store(relsPath, append([]byte(xml.Header), selfCloseEmptyElements(relsOutput)...))
	}

	// update [Content_Types].xml
	_ = f.setContentTypes("/"+layoutPath, ContentTypeSlideLayoutML)
	_ = f.setContentTypes("/"+relsPath, ContentTypeRelationships)

	// update the rels of the slide master: the relationship of the new
	// layout
	masterRelsPath := f.getSlideRelsPath(source.master)
	rID := f.addRels(masterRelsPath, SourceRelationshipSlideLayout,
		"../slideLayouts/slideLayout"+strconv.Itoa(index)+".xml", "")

	// update the slide layout id list of the slide master: the id of the
	// new layout
	if err := f.addSlideLayoutIDToMaster(source.master, nextLayoutID, rID); err != nil {
		return -1, err
	}

	return nextLayoutID, nil
}

// sldLayoutIDListEndRe matches the end of the sldLayoutIdLst element of the
// slide master.
var sldLayoutIDListEndRe = regexp.MustCompile(`</[A-Za-z0-9]+:sldLayoutIdLst>`)

// addSlideLayoutIDToMaster provides a function to add the slide layout id
// element (p:sldLayoutId) to the slide layout id list (sldLayoutIdLst) of
// the slide master: the element is inserted before the end of the list, or
// the list is created before the first of the transition, the timing, the
// header footer, the text styles and the extension list elements (the order
// of the elements of the slide master by the schema), or before the end of
// the slide master root element.
func (f *File) addSlideLayoutIDToMaster(path string, layoutID, rID int) error {
	raw := f.readBytes(path)
	prefix := "p:"
	if m := masterRootPrefixRe.FindSubmatch(raw); m != nil {
		prefix = string(m[1]) + ":"
	}
	entry := "<" + prefix + `sldLayoutId id="` + strconv.Itoa(layoutID) + `" r:id="rId` +
		strconv.Itoa(rID) + `"/>`

	if end := sldLayoutIDListEndRe.Find(raw); end != nil {
		// the entry is inserted before the end of the list
		pos := bytes.Index(raw, end)
		raw = append(append(append([]byte{}, raw[:pos]...), []byte(entry)...), raw[pos:]...)
		f.Pkg.Store(path, raw)

		return nil
	}

	// the list is absent: it is created with the single entry
	insertion := "<" + prefix + `sldLayoutIdLst>` + entry + "</" + prefix + `sldLayoutIdLst>`
	pos := len(raw)
	for _, el := range []string{"transition", "timing", "hf", "txStyles", "extLst"} {
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
	f.Pkg.Store(path, raw)

	return nil
}
