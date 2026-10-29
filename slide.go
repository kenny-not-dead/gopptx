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
	"io"
	"path/filepath"
	"strconv"
	"strings"
)

// NewSlide provides the function to create a new slide and
// returns the index of the slide in the presentation after it appended.
func (f *File) NewSlide() (int, error) {
	presentation, err := f.presentationReader()
	if err != nil {
		return -1, err
	}

	f.SlideCount++

	slideID := defaultXMLSlideID

	for _, s := range presentation.Slides.Slide {
		if s.SlideID >= slideID {
			slideID = s.SlideID + 1
		}
	}

	nextFileIndex := len(presentation.Slides.Slide) + 1
	fileName := "slide" + strconv.Itoa(nextFileIndex)

	// Update [Content_Types].xml
	_ = f.setContentTypes("/ppt/slides/_rels/"+fileName+".xml.rels", ContentTypeRelationships)
	_ = f.setContentTypes("/ppt/slides/"+fileName+".xml", ContentTypeSlideML)

	// Update presentation.xml.rels
	rID := f.addRels(f.getPresentationRelsPath(), SourceRelationshipSlide, fmt.Sprintf("slides/%s.xml", fileName), "")

	// Create new slide /ppt/slides/slide%d.xml and slide rels /ppt/slides/_rels/slide%d.xml.rels
	_ = f.setSlide(nextFileIndex, slideID)

	// Update presentation.xml
	f.setPresentation(slideID, rID)

	return slideID, nil
}

// setContentTypes provides a function to read and update property of contents
// type of the presentation.
func (f *File) setContentTypes(partName, contentType string) error {
	content, err := f.contentTypesReader()
	if err != nil {
		return err
	}
	content.mu.Lock()
	defer content.mu.Unlock()
	content.Overrides = append(content.Overrides, contentTypeOverride{
		PartName:    partName,
		ContentType: contentType,
	})
	return err
}

// contentTypesReader provides a function to get the pointer to the
// [Content_Types].xml structure after deserialization.
func (f *File) contentTypesReader() (*contentTypes, error) {
	if f.ContentTypes == nil {
		f.ContentTypes = new(contentTypes)
		f.ContentTypes.mu.Lock()
		defer f.ContentTypes.mu.Unlock()
		if err := f.xmlNewDecoder(bytes.NewReader(namespaceStrictToTransitional(f.readXML(defaultXMLPathContentTypes)))).
			Decode(f.ContentTypes); err != nil && err != io.EOF {
			return f.ContentTypes, err
		}
		// remember the default extensions declared in the origin content
		// types part, the library must not remove them on image deletion
		f.originalContentTypeExtensions = make([]string, 0, len(f.ContentTypes.Defaults))
		for _, d := range f.ContentTypes.Defaults {
			f.originalContentTypeExtensions = append(f.originalContentTypeExtensions, d.Extension)
		}
	}
	return f.ContentTypes, nil
}

// setSlide provides a function to update slide property by given index.
func (f *File) setSlide(index int, id int) error {
	slideXMLPath := "ppt/slides/slide" + strconv.Itoa(index) + ".xml"
	f.slideMap[id] = slideXMLPath

	slide := decodeSlide{}
	err := xml.Unmarshal([]byte(xml.Header+TemplateSlide), &slide)
	if err != nil {
		return err
	}

	f.Slide.Store(slideXMLPath, &slide)
	// capture the origin root element attributes of the slide template, so
	// the new slide keeps the template namespaces on save
	f.xmlAttr.Store(slideXMLPath, getRootElement(f.xmlNewDecoder(bytes.NewReader([]byte(TemplateSlide)))))

	relsSlideXMLPath := "ppt/slides/_rels/slide" + strconv.Itoa(index) + ".xml.rels"
	f.Pkg.Store(relsSlideXMLPath, []byte(xml.Header+TemplateSlideRels))

	return nil
}

// getSlideMap provides a function to get slide name and XML file path map
// of the presentation.
func (f *File) getSlideMap() (map[int]string, error) {
	maps := map[int]string{}
	presentation, err := f.presentationReader()
	if err != nil {
		return nil, err
	}
	rels, err := f.relsReader(f.getPresentationRelsPath())
	if err != nil {
		return nil, err
	}
	if rels == nil {
		return maps, nil
	}
	for _, slide := range presentation.Slides.Slide {
		for _, rel := range rels.Relationships {
			if rel.ID == slide.RelationshipID {
				slideXMLPath := f.getSlidePath(rel.Target)
				if _, ok := f.Pkg.Load(slideXMLPath); ok {
					maps[slide.SlideID] = slideXMLPath
				}
				if _, ok := f.tempFiles.Load(slideXMLPath); ok {
					maps[slide.SlideID] = slideXMLPath
				}
			}
		}
	}

	return maps, nil
}

// relsReader provides a function to get the pointer to the structure
// after deserialization of relationships parts.
func (f *File) relsReader(path string) (*relationships, error) {
	rels, _ := f.Relationships.Load(path)
	if rels == nil {
		if _, ok := f.Pkg.Load(path); ok {
			c := relationships{}
			if err := f.xmlNewDecoder(bytes.NewReader(namespaceStrictToTransitional(f.readXML(path)))).
				Decode(&c); err != nil && err != io.EOF {
				return nil, err
			}
			f.Relationships.Store(path, &c)
		}
	}
	if rels, _ = f.Relationships.Load(path); rels != nil {
		return rels.(*relationships), nil
	}
	return nil, nil
}

// getSlidePath construct a target XML as ppt/slides/slide%d by split
// path, compatible with different types of relative paths in
// presentation.xml.rels, for example: slides/slide%d.xml
// and /ppt/slides/slide%d.xml
func (f *File) getSlidePath(relTarget string) (path string) {
	path = filepath.ToSlash(strings.TrimPrefix(
		strings.ReplaceAll(filepath.Clean(fmt.Sprintf("%s/%s", filepath.Dir(f.getPresentationPath()), relTarget)), "\\", "/"), "/"))
	if strings.HasPrefix(relTarget, "/") {
		path = filepath.ToSlash(strings.TrimPrefix(strings.ReplaceAll(filepath.Clean(relTarget), "\\", "/"), "/"))
	}
	return path
}

// getSlideXMLPath provides a function to get XML file path by given slide id.
func (f *File) getSlideXMLPath(id int) (string, bool) {
	path, ok := f.slideMap[id]
	return path, ok
}

// GetShapes provides a function to get shapes by given slide id, in the
// document order. The shapes are returned by value: edits of the returned
// slice do not affect the slide, use the SetShape* methods to change a shape.
func (f *File) GetShapes(slideID int) ([]decodeShape, error) {
	var shapes []decodeShape
	slide, err := f.slideReader(slideID)
	if err != nil {
		return shapes, err
	}

	return slide.getShapes(), err
}

// getShapes returns the shapes of the slide, in the document order.
func (ds *decodeSlide) getShapes() []decodeShape {
	var shapes []decodeShape
	for i := range ds.CommonSlideData.ShapeTree.Children {
		if child := ds.CommonSlideData.ShapeTree.Children[i].Shape; child != nil {
			shapes = append(shapes, *child)
		}
	}

	return shapes
}

// SetShapeTextBody provides a function to set shape text body by given shape id.
func (f *File) SetShapeTextBody(slideID int, shapeID int, textBody DecodeTextBody) error {
	slide, err := f.slideReader(slideID)
	if err != nil {
		return err
	}

	for i := range slide.CommonSlideData.ShapeTree.Children {
		shape := slide.CommonSlideData.ShapeTree.Children[i].Shape
		if shape == nil ||
			shape.NonVisualShapeProperties == nil ||
			shape.NonVisualShapeProperties.CommonNonVisualProperties == nil ||
			shape.NonVisualShapeProperties.CommonNonVisualProperties.ID != shapeID {
			continue
		}

		shape.TextBody = &textBody

		return nil
	}

	return ErrShapeNotExist{shapeID}
}

// SetSlideBackground provides a function to set the slide background by given
// slide id, replacing the current background of the slide.
func (f *File) SetSlideBackground(slideID int, background DecodeBackground) error {
	slide, err := f.slideReader(slideID)
	if err != nil {
		return err
	}

	slide.CommonSlideData.Background = &background

	return nil
}

// GetSlideBackground provides a function to get the slide background by given
// slide id. It returns nil when the slide has no explicit background and
// inherits it from the slide layout or the slide master.
func (f *File) GetSlideBackground(slideID int) (*DecodeBackground, error) {
	slide, err := f.slideReader(slideID)
	if err != nil {
		return nil, err
	}

	return slide.CommonSlideData.Background, nil
}

// DeleteSlideBackground provides a function to delete the slide background by
// given slide id, so the slide inherits the background from the slide layout
// or the slide master.
func (f *File) DeleteSlideBackground(slideID int) error {
	slide, err := f.slideReader(slideID)
	if err != nil {
		return err
	}

	slide.CommonSlideData.Background = nil

	return nil
}

// SetShapeFill provides a function to set the solid fill of the shape by
// given shape id, replacing the current fill of the shape.
func (f *File) SetShapeFill(slideID int, shapeID int, fill DecodeSolidFill) error {
	slide, err := f.slideReader(slideID)
	if err != nil {
		return err
	}

	for i := range slide.CommonSlideData.ShapeTree.Children {
		shape := slide.CommonSlideData.ShapeTree.Children[i].Shape
		if shape == nil ||
			shape.NonVisualShapeProperties == nil ||
			shape.NonVisualShapeProperties.CommonNonVisualProperties == nil ||
			shape.NonVisualShapeProperties.CommonNonVisualProperties.ID != shapeID {
			continue
		}
		if shape.ShapeProperties == nil {
			shape.ShapeProperties = &DecodeShapeProperties{}
		}
		shape.ShapeProperties.NoFill = nil
		shape.ShapeProperties.SolidFill = &fill

		return nil
	}

	return ErrShapeNotExist{shapeID}
}

// GetShapeFill provides a function to get the solid fill of the shape by
// given shape id. It returns nil when the shape has no solid fill.
func (f *File) GetShapeFill(slideID int, shapeID int) (*DecodeSolidFill, error) {
	slide, err := f.slideReader(slideID)
	if err != nil {
		return nil, err
	}

	for i := range slide.CommonSlideData.ShapeTree.Children {
		shape := slide.CommonSlideData.ShapeTree.Children[i].Shape
		if shape == nil ||
			shape.NonVisualShapeProperties == nil ||
			shape.NonVisualShapeProperties.CommonNonVisualProperties == nil ||
			shape.NonVisualShapeProperties.CommonNonVisualProperties.ID != shapeID {
			continue
		}
		if shape.ShapeProperties == nil {
			return nil, nil
		}

		return shape.ShapeProperties.SolidFill, nil
	}

	return nil, ErrShapeNotExist{shapeID}
}

// DeleteShapeFill provides a function to delete the fill of the shape by
// given shape id, setting the noFill property of the shape instead.
func (f *File) DeleteShapeFill(slideID int, shapeID int) error {
	slide, err := f.slideReader(slideID)
	if err != nil {
		return err
	}

	for i := range slide.CommonSlideData.ShapeTree.Children {
		shape := slide.CommonSlideData.ShapeTree.Children[i].Shape
		if shape == nil ||
			shape.NonVisualShapeProperties == nil ||
			shape.NonVisualShapeProperties.CommonNonVisualProperties == nil ||
			shape.NonVisualShapeProperties.CommonNonVisualProperties.ID != shapeID {
			continue
		}
		if shape.ShapeProperties == nil {
			shape.ShapeProperties = &DecodeShapeProperties{}
		}
		shape.ShapeProperties.NoFill = &noFill{}
		shape.ShapeProperties.SolidFill = nil

		return nil
	}

	return ErrShapeNotExist{shapeID}
}

// SetShapeXfrm provides a function to set the position and the size
// (a:xfrm) of the shape by given slide id and shape id.
func (f *File) SetShapeXfrm(slideID int, shapeID int, xfrm DecodeXfrm) error {
	slide, err := f.slideReader(slideID)
	if err != nil {
		return err
	}

	for i := range slide.CommonSlideData.ShapeTree.Children {
		shape := slide.CommonSlideData.ShapeTree.Children[i].Shape
		if shape == nil ||
			shape.NonVisualShapeProperties == nil ||
			shape.NonVisualShapeProperties.CommonNonVisualProperties == nil ||
			shape.NonVisualShapeProperties.CommonNonVisualProperties.ID != shapeID {
			continue
		}
		if shape.ShapeProperties == nil {
			shape.ShapeProperties = &DecodeShapeProperties{}
		}
		shape.ShapeProperties.Xfrm = &xfrm

		return nil
	}

	return ErrShapeNotExist{shapeID}
}

// ShapeOrder is a single entry of the order (the z-order) of the children
// of the slide shape tree: the kind of the child and its id (the id
// attribute of the p:cNvPr element of the child).
type ShapeOrder struct {
	Kind string
	ID   int
}

// The kinds of the children of the slide shape tree, the local names of
// the elements in the slide XML.
const (
	// ShapeOrderSp is the kind of the p:sp child of the shape tree.
	ShapeOrderSp = "sp"
	// ShapeOrderPic is the kind of the p:pic child of the shape tree.
	ShapeOrderPic = "pic"
	// ShapeOrderCxnSp is the kind of the p:cxnSp child of the shape tree.
	ShapeOrderCxnSp = "cxnSp"
)

// GetShapeOrder provides a function to get the order (the z-order) of the
// children of the shape tree of the slide by given slide id.
func (f *File) GetShapeOrder(slideID int) ([]ShapeOrder, error) {
	var shapeOrder []ShapeOrder
	slide, err := f.slideReader(slideID)
	if err != nil {
		return shapeOrder, err
	}

	for i := range slide.CommonSlideData.ShapeTree.Children {
		if entry, ok := shapeOrderOf(&slide.CommonSlideData.ShapeTree.Children[i]); ok {
			shapeOrder = append(shapeOrder, entry)
		}
	}

	return shapeOrder, nil
}

// SetShapeOrder provides a function to set the order (the z-order) of the
// children of the shape tree of the slide by given slide id: the given
// order must contain every child of the shape tree of the slide exactly
// once.
func (f *File) SetShapeOrder(slideID int, shapeOrder []ShapeOrder) error {
	slide, err := f.slideReader(slideID)
	if err != nil {
		return err
	}

	children := slide.CommonSlideData.ShapeTree.Children

	// the given order must contain every child of the shape tree exactly
	// once
	index := make(map[ShapeOrder]int, len(children))
	for i := range children {
		entry, ok := shapeOrderOf(&children[i])
		if !ok {
			// the child has no id (the cNvPr element is absent in the
			// slide XML): it can not be identified in the given order
			return ErrShapeOrderInvalid
		}
		index[entry] = i
	}
	if len(shapeOrder) != len(index) {
		return ErrShapeOrderInvalid
	}
	seen := make(map[ShapeOrder]bool, len(shapeOrder))
	for _, entry := range shapeOrder {
		if _, ok := index[entry]; !ok || seen[entry] {
			return ErrShapeOrderInvalid
		}
		seen[entry] = true
	}

	reordered := make([]decodeShapeTreeChild, 0, len(shapeOrder))
	for _, entry := range shapeOrder {
		reordered = append(reordered, children[index[entry]])
	}
	slide.CommonSlideData.ShapeTree.Children = reordered

	return nil
}

// MoveShape moves the shape to a specified position in the shape tree of
// the slide: the shape with the given source id is moved before the shape
// with the given target id (the same way MoveSheet of the excelize library
// moves the sheet before the target sheet).
func (f *File) MoveShape(slideID int, sourceID, targetID int) error {
	shapeOrder, err := f.GetShapeOrder(slideID)
	if err != nil {
		return err
	}

	sourceIdx, targetIdx := -1, -1
	for i, entry := range shapeOrder {
		if entry.ID == sourceID {
			sourceIdx = i
		}
		if entry.ID == targetID {
			targetIdx = i
		}
	}
	if sourceIdx == -1 {
		return ErrShapeNotExist{sourceID}
	}
	if targetIdx == -1 {
		return ErrShapeNotExist{targetID}
	}
	if sourceIdx == targetIdx {
		return nil
	}

	moved := shapeOrder[sourceIdx]
	shapeOrder = append(shapeOrder[:sourceIdx], shapeOrder[sourceIdx+1:]...)
	if targetIdx > sourceIdx {
		targetIdx--
	}
	shapeOrder = append(shapeOrder[:targetIdx], append([]ShapeOrder{moved}, shapeOrder[targetIdx:]...)...)

	return f.SetShapeOrder(slideID, shapeOrder)
}

// shapeOrderOf returns the order entry (the kind and the id) of the child
// of the shape tree. It returns false when the child has no common non-
// visual properties (the cNvPr element is absent in the slide XML).
func shapeOrderOf(child *decodeShapeTreeChild) (ShapeOrder, bool) {
	switch {
	case child.Shape != nil:
		if child.Shape.NonVisualShapeProperties == nil ||
			child.Shape.NonVisualShapeProperties.CommonNonVisualProperties == nil {
			return ShapeOrder{}, false
		}

		return ShapeOrder{
			Kind: ShapeOrderSp,
			ID:   child.Shape.NonVisualShapeProperties.CommonNonVisualProperties.ID,
		}, true
	case child.Picture != nil:
		if child.Picture.NonVisualPictureProperties == nil ||
			child.Picture.NonVisualPictureProperties.CommonNonVisualProperties == nil {
			return ShapeOrder{}, false
		}

		return ShapeOrder{
			Kind: ShapeOrderPic,
			ID:   child.Picture.NonVisualPictureProperties.CommonNonVisualProperties.ID,
		}, true
	case child.Connector != nil:
		if child.Connector.NonVisualConnectorProperties == nil ||
			child.Connector.NonVisualConnectorProperties.CommonNonVisualProperties == nil {
			return ShapeOrder{}, false
		}

		return ShapeOrder{
			Kind: ShapeOrderCxnSp,
			ID:   child.Connector.NonVisualConnectorProperties.CommonNonVisualProperties.ID,
		}, true
	}

	return ShapeOrder{}, false
}

// CreateShape provides the function to create  a new shape by given slide id and
// returns the id of the shape in the slide after it appended.
func (f *File) CreateShape(slideID int, shapeProperties DecodeShapeProperties, textBody DecodeTextBody) (int, error) {
	slide, err := f.slideReader(slideID)
	if err != nil {
		return -1, err
	}

	// the shape id must be unique in the whole shape tree (sp + pic +
	// cxnSp), the same way AddImage and CreateConnector allocate shape ids
	shapeID := f.nextShapeID(slide)

	// the body properties element is required in the text body by the schema
	if textBody.BodyProperties == nil {
		textBody.BodyProperties = &DecodeBodyProperties{}
	}

	newShape := decodeShape{
		NonVisualShapeProperties: &decodeNonVisualShapeProperties{
			CommonNonVisualProperties: &decodeCommonNonVisualProperties{
				ID: shapeID,
			},
			// the shape non-visual properties and the non-visual properties
			// elements are expected by the consumers (PowerPoint writes them
			// for every shape, python-pptx requires the former)
			CommonNonVisualShapeProperties: &decodeCommonNonVisualShapeProperties{},
			NonVisualProperties:            &decodeNonVisualProperties{},
		},
		ShapeProperties: &shapeProperties,
		TextBody:        &textBody,
	}

	// the new shape is appended to the end of the shape tree: it is drawn
	// above the existing shapes of the slide
	slide.CommonSlideData.ShapeTree.Children = append(slide.CommonSlideData.ShapeTree.Children, decodeShapeTreeChild{Shape: &newShape})

	return shapeID, nil
}

// DeleteShape provides a function to delete shape on slide by given shape id.
func (f *File) DeleteShape(slideID int, shapeID int) error {
	slide, err := f.slideReader(slideID)
	if err != nil {
		return err
	}

	children := slide.CommonSlideData.ShapeTree.Children

	deleteSlideIndex := -1
	for i := range children {
		if s := children[i].Shape; s != nil &&
			s.NonVisualShapeProperties != nil &&
			s.NonVisualShapeProperties.CommonNonVisualProperties != nil &&
			s.NonVisualShapeProperties.CommonNonVisualProperties.ID == shapeID {

			deleteSlideIndex = i

			break
		}
	}

	if deleteSlideIndex == -1 {
		return ErrShapeNotExist{shapeID}
	}

	slide.CommonSlideData.ShapeTree.Children = append(children[:deleteSlideIndex], children[deleteSlideIndex+1:]...)

	return nil
}

// GetGroupShapeProperties provides a function to get group shape properties by given slide id.
func (f *File) GetGroupShapeProperties(slideID int) (*decodeGroupShapeProperties, error) {
	s, err := f.slideReader(slideID)
	if err != nil {
		return nil, err
	}

	return s.getGroupShapeProperties(), err
}

// getGroupShapeProperties returns group shape properties of the slide.
func (ds *decodeSlide) getGroupShapeProperties() *decodeGroupShapeProperties {
	return ds.CommonSlideData.ShapeTree.GroupShapeProperties
}

// GetNonVisualGroupShapeProperties provides a function to get non visual group shape properties by given slide id.
func (f *File) GetNonVisualGroupShapeProperties(slideID int) (*decodeNonVisualGroupShapeProperties, error) {
	s, err := f.slideReader(slideID)
	if err != nil {
		return nil, err
	}

	return s.getNonVisualGroupShapeProperties(), err
}

// NonVisualGroupShapeProperties returns non visual group shape properties of the slide.
func (ds *decodeSlide) getNonVisualGroupShapeProperties() *decodeNonVisualGroupShapeProperties {
	return ds.CommonSlideData.ShapeTree.NonVisualGroupShapeProperties
}

// DeleteSlide provides a function to delete slide in a presentation by given slide id.
func (f *File) DeleteSlide(slideID int) error {
	if idx, _ := f.GetSlideIndex(slideID); f.SlideCount == 1 || idx == -1 {
		return nil
	}

	presentation, _ := f.presentationReader()
	presentationRels, _ := f.relsReader(f.getPresentationRelsPath())

	for idx, v := range presentation.Slides.Slide {
		if v.SlideID != slideID {
			continue
		}

		presentation.Slides.Slide = append(presentation.Slides.Slide[:idx], presentation.Slides.Slide[idx+1:]...)
		var slideXML, rels string
		if presentationRels != nil {
			for _, rel := range presentationRels.Relationships {
				if rel.ID == v.RelationshipID {
					slideXML = f.getSlidePath(rel.Target)
					rels = "ppt/slides/_rels/" + strings.TrimPrefix(slideXML, "ppt/slides/") + ".rels"
				}
			}
		}

		target := f.deleteSlideFromPresentationRels(v.RelationshipID)
		dir := filepath.Dir(target)
		base := filepath.Base(target)

		// Update [Content_Types].xml
		_ = f.removeContentTypesPart(ContentTypeSlideML, target)
		_ = f.removeContentTypesPart(ContentTypeRelationships, filepath.Join(dir, "_rels", base+".rels"))

		// delete the notes slide of the slide (the part, the rels part and
		// the content types; the notesSlide relationship of the slide is
		// removed together with the slide rels part)
		if notesPath, ok := f.getNotesSlidePath(v.SlideID); ok {
			notesBase := filepath.Base(notesPath)
			notesRelsPath := "ppt/notesSlides/_rels/" + notesBase + ".rels"
			_ = f.removeContentTypesPart(ContentTypeNotesSlideML, notesPath)
			_ = f.removeContentTypesPart(ContentTypeRelationships, "ppt/notesSlides/_rels/"+notesBase+".rels")
			f.Pkg.Delete(notesPath)
			f.Pkg.Delete(notesRelsPath)
			f.Relationships.Delete(notesRelsPath)
			f.NotesSlides.Delete(notesPath)
			f.xmlAttr.Delete(notesPath)
		}

		delete(f.slideMap, v.SlideID)
		f.Pkg.Delete(slideXML)
		f.Pkg.Delete(rels)
		f.Relationships.Delete(rels)
		f.Slide.Delete(slideXML)
		f.xmlAttr.Delete(slideXML)
		f.SlideCount--
	}

	// TODO: setActiveSlide
	//index, err := f.GetSlideIndex(f.getActiveSlideID())
	//f.SetActiveSlide(index)
	return nil
}

// deleteSlideFromPresentationRels provides a function to remove slide
// relationships by given relationships ID in the file presentation.xml.rels.
func (f *File) deleteSlideFromPresentationRels(rID string) string {
	rels, _ := f.relsReader(f.getPresentationRelsPath())
	rels.mu.Lock()
	defer rels.mu.Unlock()
	for k, v := range rels.Relationships {
		if v.ID == rID {
			rels.Relationships = append(rels.Relationships[:k], rels.Relationships[k+1:]...)
			return v.Target
		}
	}
	return ""
}

// GetSlideIndex provides a function to get a slide index of the presentation by
// the given slide id. If slide doesn't exist, it will return an integer type value -1.
func (f *File) GetSlideIndex(slideID int) (int, error) {
	for index, id := range f.GetSlideList() {
		if id == slideID {
			return index, nil
		}
	}
	return -1, nil
}

// GetSlideList provides a function to get slides of the presentation.
func (f *File) GetSlideList() (list []int) {
	presentation, _ := f.presentationReader()
	if presentation != nil {
		for _, slide := range presentation.Slides.Slide {
			list = append(list, slide.SlideID)
		}
	}
	return
}

// GetActiveSlideIndex provides a function to get active slide index of the
// presentation. If not found the active slide will be return integer 0.
func (f *File) GetActiveSlideIndex() (index int) {
	slideID := f.getActiveSlideID()
	presentation, _ := f.presentationReader()
	if presentation != nil {
		for idx, slide := range presentation.Slides.Slide {
			if slide.SlideID == slideID {
				index = idx
				return
			}
		}
	}
	return
}

// getActiveSlideID provides a function to get active slide ID of the
// presentation. If not found the active slide will be return integer 0.
func (f *File) getActiveSlideID() int {
	presentation, _ := f.presentationReader()
	if presentation != nil {
		// TODO: get active slide
		if len(presentation.Slides.Slide) >= 1 {
			return presentation.Slides.Slide[0].SlideID
		}
	}
	return 0
}
