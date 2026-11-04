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
)

// slideMasterRef is a slide master of the presentation: the id of the master
// (the id attribute of the p:sldMasterId element of the presentation) and the
// part path of the slide master.
type slideMasterRef struct {
	id     int
	master string
}

// slideMasterRefs provides a function to get the slide masters of the
// presentation (the p:sldMasterId elements of the sldMasterIdLst element of
// the presentation), in the document order.
func (f *File) slideMasterRefs() ([]slideMasterRef, error) {
	presentation, err := f.presentationReader()
	if err != nil {
		return nil, err
	}
	rels, err := f.relsReader(f.getPresentationRelsPath())
	if err != nil {
		return nil, err
	}
	var refs []slideMasterRef
	for _, master := range presentation.MasterSlide.MasterSlide {
		if master.RelationshipID == "" {
			continue
		}
		for _, rel := range rels.Relationships {
			if rel.ID != master.RelationshipID {
				continue
			}
			refs = append(refs, slideMasterRef{
				id:     master.SlideID,
				master: resolvePartPath(f.getPresentationPath(), rel.Target),
			})
		}
	}

	return refs, nil
}

// masterReader provides a function to get the parsed slide master by the
// master id (the same way the slideReader gets the slide by the slide id),
// and the part path of the slide master.
func (f *File) masterReader(masterID int) (*decodeSlideMaster, string, error) {
	refs, err := f.slideMasterRefs()
	if err != nil {
		return nil, "", err
	}
	for _, ref := range refs {
		if ref.id != masterID {
			continue
		}
		dm := new(decodeSlideMaster)
		if err := f.xmlNewDecoder(bytes.NewReader(f.readBytes(ref.master))).Decode(dm); err != nil && err != io.EOF {
			return nil, "", err
		}

		return dm, ref.master, nil
	}

	return nil, "", ErrSlideMasterNotExist{masterID}
}

// masterCheckModifiable provides a function to check the slide master for
// the elements that are not supported by the master modification (the
// modification of the master is not supported for such parts, so the master
// cannot be modified without losing the content of such elements). The
// elements of the slide master (the color map, the slide layout id list, the
// timing, the header footer, the text styles and the extension list) are
// preserved as-is on the master round trip.
func (f *File) masterCheckModifiable(path string, dm *decodeSlideMaster) error {
	raw := f.readBytes(path)
	if element := commonSlideDataUnsupportedElement(raw); element != "" {
		return ErrSlideMasterUnsupported{Master: path, Element: element}
	}
	// the top level elements of the slide master that are not modeled by the
	// master round trip (the markup compatibility elements of the slide master
	// root and the transition element, the attributes of the transition
	// element are not modeled, the same as in the slide layout round trip)
	if i := bytes.Index(raw, []byte("</p:cSld>")); i != -1 {
		if bytes.Contains(raw[i:], []byte("<mc:AlternateContent")) {
			return ErrSlideMasterUnsupported{
				Master:  path,
				Element: "the markup compatibility element of the slide master",
			}
		}
		if bytes.Contains(raw[i:], []byte(":transition")) {
			return ErrSlideMasterUnsupported{
				Master:  path,
				Element: "the transition element of the slide master",
			}
		}
	}
	for i := range dm.CommonSlideData.ShapeTree.Children {
		child := &dm.CommonSlideData.ShapeTree.Children[i]
		if child.Shape == nil {
			return ErrSlideMasterUnsupported{Master: path, Element: "the shape of the type other than sp"}
		}
	}

	return nil
}

// saveSlideMaster provides a function to save the slide master part after
// the serialize structure (the same way the slideWriter saves the slide
// part). The root element of the part keeps its origin attributes (the
// namespaces and the attributes), the elements of the master other than the
// cSld element are written from the origin part as-is.
func (f *File) saveSlideMaster(path string, dm *decodeSlideMaster) error {
	if _, ok := f.xmlAttr.Load(path); !ok {
		f.xmlAttr.Store(path, getRootElement(f.xmlNewDecoder(bytes.NewReader(f.readBytes(path)))))
	}

	var (
		shapes []Shape
		order  []shapeTreeOrderEntry
	)
	for i := range dm.CommonSlideData.ShapeTree.Children {
		child := &dm.CommonSlideData.ShapeTree.Children[i]
		if child.Shape != nil {
			shapes = append(shapes, newShape(*child.Shape))
			order = append(order, shapeTreeOrderEntry{kind: shapeTreeChildSp, index: len(shapes) - 1})
		}
	}

	output, err := xml.Marshal(&slideMasterPart{
		XMLNSA:   NameSpaceDrawingML.Value,
		XMLNSP:   NameSpacePresentationML.Value,
		XMLNSR:   SourceRelationship.Value,
		XMLNSP14: NameSpacePowerPointR14.Value,
		XMLNSP15: NameSpacePowerPointR15.Value,
		XMLNSMC:  SourceRelationshipCompatibility.Value,
		CommonSlideData: SlideData{
			Name:       dm.CommonSlideData.Name,
			Background: newBG(dm.CommonSlideData.Background),
			ShapeTree: ShapeTree{
				NonVisualGroupShapeProperties: newNonVisualGroupShapeProperties(dm.CommonSlideData.ShapeTree.NonVisualGroupShapeProperties),
				GroupShapeProperties:          newGroupShapeProperties(dm.CommonSlideData.ShapeTree.GroupShapeProperties),
				Shape:                         shapes,
				shapeOrder:                    order,
			},
			ExtLst: newSlideDataExtLst(dm.CommonSlideData.ExtLst),
		},
		ColorMap:          dm.ColorMap,
		SlideLayoutIDList: dm.SlideLayoutIDList,
		Transition:        dm.Transition,
		Timing:            dm.Timing,
		HeaderFooter:      dm.HeaderFooter,
		TextStyles:        dm.TextStyles,
		ExtLst:            dm.ExtLst,
	})
	if err != nil {
		return err
	}

	f.saveFileList(path, f.replaceNameSpaceBytes(path, output))
	return nil
}

// GetSlideMasterList provides a function to get the id list of the slide
// masters of the presentation (the id attributes of the p:sldMasterId
// elements of the presentation).
func (f *File) GetSlideMasterList() (list []int) {
	refs, err := f.slideMasterRefs()
	if err != nil {
		return nil
	}
	for _, ref := range refs {
		list = append(list, ref.id)
	}

	return list
}

// GetMasterShapes provides a function to get the shapes of the slide master
// by the given master id, in the document order.
func (f *File) GetMasterShapes(masterID int) ([]decodeShape, error) {
	dm, _, err := f.masterReader(masterID)
	if err != nil {
		return nil, err
	}

	return dm.CommonSlideData.shapes(), nil
}

// CreateMasterShape provides a function to create a new shape (p:sp) in the
// shape tree of the slide master by the given master id, and returns the id
// of the new shape. The new shape is appended to the end of the shape tree:
// it is drawn above the existing shapes of the master and is inherited by
// all the slides of the master that do not hide the master shapes.
func (f *File) CreateMasterShape(masterID int, shapeProperties DecodeShapeProperties, textBody DecodeTextBody) (int, error) {
	dm, path, err := f.masterReader(masterID)
	if err != nil {
		return -1, err
	}
	if err := f.masterCheckModifiable(path, dm); err != nil {
		return -1, err
	}

	shapeID := dm.CommonSlideData.nextShapeID()
	shape := newDecodeShape(shapeID, shapeProperties, textBody)
	dm.CommonSlideData.ShapeTree.Children = append(
		dm.CommonSlideData.ShapeTree.Children, decodeShapeTreeChild{Shape: &shape})

	return shapeID, f.saveSlideMaster(path, dm)
}

// DeleteMasterShape provides a function to delete the shape by the given
// shape id from the shape tree of the slide master by the given master id.
func (f *File) DeleteMasterShape(masterID int, shapeID int) error {
	dm, path, err := f.masterReader(masterID)
	if err != nil {
		return err
	}
	if err := f.masterCheckModifiable(path, dm); err != nil {
		return err
	}
	if !dm.CommonSlideData.deleteShape(shapeID) {
		return ErrShapeNotExist{shapeID}
	}

	return f.saveSlideMaster(path, dm)
}

// SetMasterShapeTextBody provides a function to set the shape text body by
// the given master id and shape id, replacing the current text body of the
// shape of the slide master.
func (f *File) SetMasterShapeTextBody(masterID int, shapeID int, textBody DecodeTextBody) error {
	dm, path, err := f.masterReader(masterID)
	if err != nil {
		return err
	}
	if err := f.masterCheckModifiable(path, dm); err != nil {
		return err
	}
	if shape := dm.CommonSlideData.findShape(shapeID); shape == nil {
		return ErrShapeNotExist{shapeID}
	} else {
		shape.TextBody = &textBody
	}

	return f.saveSlideMaster(path, dm)
}

// SetMasterShapeText provides a function to set the text of the shape of the
// slide master by the given master id and shape id, replacing the paragraph
// of the text body of the shape (the same way the SetShapeText function sets
// the text of the slide shape). The body properties of the current text body
// are kept.
func (f *File) SetMasterShapeText(masterID int, shapeID int, text string) error {
	dm, path, err := f.masterReader(masterID)
	if err != nil {
		return err
	}
	if err := f.masterCheckModifiable(path, dm); err != nil {
		return err
	}
	shape := dm.CommonSlideData.findShape(shapeID)
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

	return f.saveSlideMaster(path, dm)
}

// SetMasterShapeXfrm provides a function to set the position and the size
// (a:xfrm) of the shape of the slide master by the given master id and
// shape id.
func (f *File) SetMasterShapeXfrm(masterID int, shapeID int, xfrm DecodeXfrm) error {
	dm, path, err := f.masterReader(masterID)
	if err != nil {
		return err
	}
	if err := f.masterCheckModifiable(path, dm); err != nil {
		return err
	}
	shape := dm.CommonSlideData.findShape(shapeID)
	if shape == nil {
		return ErrShapeNotExist{shapeID}
	}
	if shape.ShapeProperties == nil {
		shape.ShapeProperties = &DecodeShapeProperties{}
	}
	shape.ShapeProperties.Xfrm = &xfrm

	return f.saveSlideMaster(path, dm)
}

// SetMasterShapeFill provides a function to set the solid fill of the shape
// of the slide master by the given master id and shape id, replacing the
// current fill of the shape.
func (f *File) SetMasterShapeFill(masterID int, shapeID int, fill DecodeSolidFill) error {
	dm, path, err := f.masterReader(masterID)
	if err != nil {
		return err
	}
	if err := f.masterCheckModifiable(path, dm); err != nil {
		return err
	}
	shape := dm.CommonSlideData.findShape(shapeID)
	if shape == nil {
		return ErrShapeNotExist{shapeID}
	}
	if shape.ShapeProperties == nil {
		shape.ShapeProperties = &DecodeShapeProperties{}
	}
	shape.ShapeProperties.NoFill = nil
	shape.ShapeProperties.SolidFill = &fill

	return f.saveSlideMaster(path, dm)
}

// GetMasterShapeFill provides a function to get the solid fill of the shape
// of the slide master by the given master id and shape id. It returns nil
// when the shape has no solid fill.
func (f *File) GetMasterShapeFill(masterID int, shapeID int) (*DecodeSolidFill, error) {
	dm, _, err := f.masterReader(masterID)
	if err != nil {
		return nil, err
	}
	shape := dm.CommonSlideData.findShape(shapeID)
	if shape == nil {
		return nil, ErrShapeNotExist{shapeID}
	}
	if shape.ShapeProperties == nil {
		return nil, nil
	}

	return shape.ShapeProperties.SolidFill, nil
}

// SetSlideMasterBackground provides a function to set the background of the
// slide master by the given master id, replacing the current background of
// the master.
func (f *File) SetSlideMasterBackground(masterID int, background DecodeBackground) error {
	dm, path, err := f.masterReader(masterID)
	if err != nil {
		return err
	}
	if err := f.masterCheckModifiable(path, dm); err != nil {
		return err
	}

	dm.CommonSlideData.Background = &background

	return f.saveSlideMaster(path, dm)
}

// GetSlideMasterBackground provides a function to get the background of the
// slide master by the given master id. It returns nil when the master has no
// explicit background.
func (f *File) GetSlideMasterBackground(masterID int) (*DecodeBackground, error) {
	dm, _, err := f.masterReader(masterID)
	if err != nil {
		return nil, err
	}

	return dm.CommonSlideData.Background, nil
}

// DeleteSlideMasterBackground provides a function to delete the background
// of the slide master by the given master id.
func (f *File) DeleteSlideMasterBackground(masterID int) error {
	dm, path, err := f.masterReader(masterID)
	if err != nil {
		return err
	}
	if err := f.masterCheckModifiable(path, dm); err != nil {
		return err
	}

	dm.CommonSlideData.Background = nil

	return f.saveSlideMaster(path, dm)
}
