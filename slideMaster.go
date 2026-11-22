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
	"path/filepath"
	"regexp"
	"strconv"
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
		TextStyles:        newTextStyles(dm.TextStyles),
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

// GetMasterShapeOrder provides a function to get the order (the z-order) of
// the children of the shape tree of the slide master by the given master id,
// in the document order.
func (f *File) GetMasterShapeOrder(masterID int) ([]ShapeOrder, error) {
	dm, _, err := f.masterReader(masterID)
	if err != nil {
		return nil, err
	}

	return dm.CommonSlideData.shapeTreeOrder(), nil
}

// SetMasterShapeOrder provides a function to set the order (the z-order) of
// the children of the shape tree of the slide master by the given master id:
// the given order must contain every child of the shape tree of the master
// exactly once. The modification of the master is not supported for the
// masters with the elements that are not modeled by the master round trip
// (the masterCheckModifiable function returns the
// ErrSlideMasterUnsupported error for such masters).
func (f *File) SetMasterShapeOrder(masterID int, shapeOrder []ShapeOrder) error {
	dm, path, err := f.masterReader(masterID)
	if err != nil {
		return err
	}
	if err := f.masterCheckModifiable(path, dm); err != nil {
		return err
	}
	if err := dm.CommonSlideData.reorderShapeTree(shapeOrder); err != nil {
		return err
	}

	return f.saveSlideMaster(path, dm)
}

// MoveMasterShape moves the shape to a specified position in the shape tree
// of the slide master by the given master id: the shape with the given
// source id is moved before the shape with the given target id (the same
// way the MoveShape function moves the shape of the slide before the target
// shape).
func (f *File) MoveMasterShape(masterID int, sourceID, targetID int) error {
	shapeOrder, err := f.GetMasterShapeOrder(masterID)
	if err != nil {
		return err
	}
	moved, err := moveShapeOrderEntry(shapeOrder, sourceID, targetID)
	if err != nil {
		return err
	}

	return f.SetMasterShapeOrder(masterID, moved)
}

// SetMasterShapeHyperLink provides a function to set the hyperlink of a
// shape of the slide master by the given master id and shape id (the same
// way the SetShapeHyperLink function sets the hyperlink of the shape of the
// slide). The linkType defines the type of the hyperlink: LinkTypeExternal
// for a URL (a hyperlink relationship with the TargetMode "External" is
// added to the rels part of the slide master), LinkTypeLocation for the
// internal link to the slide of the presentation (the link is the id of the
// target slide, a relationship of the type "slide" is added to the rels
// part of the slide master and the "ppaction://hlinksldjump" action is
// written to the hyperlink element) and LinkTypeNone to remove the
// hyperlink from the shape. The modification of the master is not supported
// for the masters with the elements that are not modeled by the master
// round trip (the masterCheckModifiable function returns the
// ErrSlideMasterUnsupported error for such masters).
func (f *File) SetMasterShapeHyperLink(masterID, shapeID int, link string, linkType LinkType, opts ...HyperlinkOpts) error {
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

	relsPath := f.getSlideRelsPath(path)
	var hlink *Hyperlink
	switch linkType {
	case LinkTypeExternal:
		hlink = f.partHyperlinkRelationship(path, relsPath, link)
	case LinkTypeLocation:
		targetSlideID, err := strconv.Atoi(link)
		if err != nil {
			return ErrLinkTargetInvalid{Target: link}
		}
		hlink, err = f.partLocationRelationship(path, relsPath, targetSlideID)
		if err != nil {
			return err
		}
	case LinkTypeNone:
		clearShapeRunHyperlinks(shape)
		f.removeUnusedLinkRels(path, shapeTreeUsedLinkIDs(dm.CommonSlideData.ShapeTree.Children))

		return f.saveSlideMaster(path, dm)
	default:
		return ErrInvalidLinkType
	}

	for _, o := range opts {
		if o.Tooltip != nil {
			hlink.Tooltip = *o.Tooltip
		}
	}
	setShapeRunHyperlinks(shape, hlink)

	return f.saveSlideMaster(path, dm)
}

// GetMasterShapeHyperLink gets the hyperlink of a shape of the slide master
// by the given master id and shape id (the same way the GetShapeHyperLink
// function gets the hyperlink of the shape of the slide). It returns true
// and the link address when the shape has a hyperlink, otherwise it returns
// false and an empty link address: the target of the hyperlink relationship
// for the URL links (LinkTypeExternal) and the id of the target slide for
// the internal links (LinkTypeLocation).
func (f *File) GetMasterShapeHyperLink(masterID, shapeID int) (bool, string, error) {
	dm, path, err := f.masterReader(masterID)
	if err != nil {
		return false, "", err
	}
	shape := dm.CommonSlideData.findShape(shapeID)
	if shape == nil {
		return false, "", ErrShapeNotExist{shapeID}
	}
	id, err := shapeHyperlinkID(shape, nil)
	if err != nil {
		return false, "", err
	}
	if id == "" {
		return false, "", nil
	}

	return f.partHyperlinkTarget(path, f.getSlideRelsPath(path), id)
}

// GetMasterShapeHyperLinks returns the shape ids that contain hyperlinks in
// the slide master by the given master id (the same way the
// GetShapeHyperLinks function returns the shape ids of the slide). The
// optional parameter linkType is used to filter the links: LinkTypeExternal
// for the URL links, LinkTypeLocation for the internal links to the parts
// of the presentation and LinkTypeNone for the shapes without hyperlinks.
// When linkType is empty, all the shapes with hyperlinks are returned.
func (f *File) GetMasterShapeHyperLinks(masterID int, linkType LinkType) ([]int, error) {
	var ids []int
	switch linkType {
	case "", LinkTypeExternal, LinkTypeLocation, LinkTypeNone:
	default:
		return ids, ErrInvalidLinkType
	}
	dm, path, err := f.masterReader(masterID)
	if err != nil {
		return ids, err
	}

	return f.shapeTreeHyperlinkIDs(dm.CommonSlideData.ShapeTree.Children, f.getSlideRelsPath(path), linkType), nil
}

// GetSlideMasterColorMap provides a function to get the color map
// (p:clrMap) of the slide master by the given master id. It returns nil
// when the master has no color map element.
func (f *File) GetSlideMasterColorMap(masterID int) (*DecodeColorMap, error) {
	dm, _, err := f.masterReader(masterID)
	if err != nil {
		return nil, err
	}

	return dm.ColorMap, nil
}

// SetSlideMasterColorMap provides a function to set the color map
// (p:clrMap) of the slide master by the given master id, replacing the
// current color map of the master. The values of the color map are the
// names of the theme scheme colors (the "dk1", the "lt1", the
// "accent1"–"accent6", the "hlink" and the "folHlink" names of the theme
// color scheme). The modification of the master is not supported for the
// masters with the elements that are not modeled by the master round trip
// (the masterCheckModifiable function returns the ErrSlideMasterUnsupported
// error for such masters).
func (f *File) SetSlideMasterColorMap(masterID int, colorMap DecodeColorMap) error {
	dm, path, err := f.masterReader(masterID)
	if err != nil {
		return err
	}
	if err := f.masterCheckModifiable(path, dm); err != nil {
		return err
	}

	dm.ColorMap = &colorMap

	return f.saveSlideMaster(path, dm)
}

// GetSlideMasterTextStyles provides a function to get the text styles
// (p:txStyles) of the slide master by the given master id. It returns nil
// when the master has no text styles element.
func (f *File) GetSlideMasterTextStyles(masterID int) (*DecodeTextStyles, error) {
	dm, _, err := f.masterReader(masterID)
	if err != nil {
		return nil, err
	}

	return dm.TextStyles, nil
}

// SetSlideMasterTextStyles provides a function to set the text styles
// (p:txStyles) of the slide master by the given master id, replacing the
// current text styles of the master. The text styles of the master are the
// default paragraph properties of the title, the body and the other text
// of the slides of the master (the styles are inherited by the slides). The
// modification of the master is not supported for the masters with the
// elements that are not modeled by the master round trip (the
// masterCheckModifiable function returns the ErrSlideMasterUnsupported
// error for such masters).
func (f *File) SetSlideMasterTextStyles(masterID int, textStyles DecodeTextStyles) error {
	dm, path, err := f.masterReader(masterID)
	if err != nil {
		return err
	}
	if err := f.masterCheckModifiable(path, dm); err != nil {
		return err
	}

	dm.TextStyles = &textStyles

	return f.saveSlideMaster(path, dm)
}

// DeleteSlideMasterTextStyles provides a function to delete the text styles
// (p:txStyles) of the slide master by the given master id.
func (f *File) DeleteSlideMasterTextStyles(masterID int) error {
	dm, path, err := f.masterReader(masterID)
	if err != nil {
		return err
	}
	if err := f.masterCheckModifiable(path, dm); err != nil {
		return err
	}

	dm.TextStyles = nil

	return f.saveSlideMaster(path, dm)
}

// slideMasterPartIndexRe matches the index of the part of the slide master
// (the N of the slideMasterN.xml part name).
var slideMasterPartIndexRe = regexp.MustCompile(`^ppt/slideMasters/slideMaster(\d+)\.xml$`)

// nextSlideMasterPartIndex provides a function to get the next free index of
// the part of the slide master (the N of the slideMasterN.xml part name).
func (f *File) nextSlideMasterPartIndex() int {
	next := 1
	f.Pkg.Range(func(k, _ interface{}) bool {
		name, _ := k.(string)
		if m := slideMasterPartIndexRe.FindStringSubmatch(name); m != nil {
			if index, _ := strconv.Atoi(m[1]); index >= next {
				next = index + 1
			}
		}
		return true
	})

	return next
}

// themePartIndexRe matches the index of the part of the theme (the N of the
// themeN.xml part name).
var themePartIndexRe = regexp.MustCompile(`^ppt/theme/theme(\d+)\.xml$`)

// nextThemePartIndex provides a function to get the next free index of the
// part of the theme (the N of the themeN.xml part name).
func (f *File) nextThemePartIndex() int {
	next := 1
	f.Pkg.Range(func(k, _ interface{}) bool {
		name, _ := k.(string)
		if m := themePartIndexRe.FindStringSubmatch(name); m != nil {
			if index, _ := strconv.Atoi(m[1]); index >= next {
				next = index + 1
			}
		}
		return true
	})

	return next
}

// NewSlideMaster provides a function to create a new slide master (a copy of
// the slide master with the given master id) in the presentation, and returns
// the id of the new master (the same way PowerPoint "duplicates" the slide
// master). The new master is the byte copy of the part of the source master:
// the shapes, the background, the color map, the text styles and the header
// footer of the source master. The slide layouts of the source master are
// copied with the new master (each layout of the new master is the copy of
// the layout of the source master and points to the new master). The theme
// of the source master is copied with the new master: the new master points
// to its own theme part (the theme part is private to the slide master in
// PowerPoint: two slide masters sharing one theme part make the repair
// prompt of PowerPoint appear). The parts of the new master and its layouts,
// the new theme part, their rels parts, the content types, the slide layout
// id list of the new master and the slide master id list of the presentation
// are updated automatically.
//
//	masterID, err := f.NewSlideMaster(f.GetSlideMasterList()[0])
func (f *File) NewSlideMaster(masterID int) (int, error) {
	refs, err := f.slideMasterRefs()
	if err != nil {
		return -1, err
	}
	var source *slideMasterRef
	nextID := 0
	for i := range refs {
		// the master ids are positive (the next id is the max id + 1)
		if refs[i].id >= nextID {
			nextID = refs[i].id + 1
		}
		if refs[i].id == masterID {
			source = &refs[i]
		}
	}
	if source == nil {
		return -1, ErrSlideMasterNotExist{masterID}
	}

	// the ids of the slide masters and the ids of the slide layouts share
	// one id space of the presentation: the new master id and the new layout
	// ids must not collide with the existing ids
	layoutRefs, err := f.slideLayoutRefs()
	if err != nil {
		return -1, err
	}
	nextLayoutID := 0
	for _, lr := range layoutRefs {
		if lr.id >= nextID {
			nextID = lr.id + 1
		}
		if lr.id >= nextLayoutID {
			nextLayoutID = lr.id + 1
		}
	}
	// the new layout ids must not collide with the new master id (one id
	// space of the presentation)
	if nextLayoutID <= nextID {
		nextLayoutID = nextID + 1
	}

	// the new master part and its rels part
	masterIndex := f.nextSlideMasterPartIndex()
	newMasterPath := "ppt/slideMasters/slideMaster" + strconv.Itoa(masterIndex) + ".xml"
	newMasterRelsPath := f.getSlideRelsPath(newMasterPath)
	newMasterBase := filepath.Base(newMasterPath)

	// the layouts of the source master (in the document order of the slide
	// layout id list of the source master) are copied with the new master
	sourceMasterRaw := f.readBytes(source.master)
	sourceMasterRels, _ := f.relsReader(f.getSlideRelsPath(source.master))
	layoutIndex := f.nextSlideLayoutPartIndex()
	type layoutCopy struct {
		oldID, newID int
		oldTarget    string // the target of the layout in the rels of the source master
		newTarget    string // the target of the new layout in the rels of the new master
	}
	var copies []layoutCopy
	for _, tag := range sldLayoutIDTagRe.FindAll(sourceMasterRaw, -1) {
		idMatch := sldLayoutIDAttrRe.FindSubmatch(tag)
		relMatch := sldLayoutIDRelAttrRe.FindSubmatch(tag)
		if idMatch == nil || relMatch == nil {
			continue
		}
		oldID, _ := strconv.Atoi(string(idMatch[1]))
		relID := string(relMatch[1])
		// the part of the layout of the source master (the target of the
		// relationship of the layout in the rels of the source master)
		var oldTarget string
		if sourceMasterRels != nil {
			sourceMasterRels.mu.Lock()
			for _, rel := range sourceMasterRels.Relationships {
				if rel.ID == relID {
					oldTarget = rel.Target
					break
				}
			}
			sourceMasterRels.mu.Unlock()
		}
		if oldTarget == "" {
			continue
		}
		oldLayoutPath := resolvePartPath(source.master, oldTarget)

		newLayoutPath := "ppt/slideLayouts/slideLayout" + strconv.Itoa(layoutIndex) + ".xml"
		layoutIndex++
		// the part of the new layout: the byte copy of the part of the
		// layout of the source master
		f.Pkg.Store(newLayoutPath, f.readBytes(oldLayoutPath))
		// the rels part of the new layout: the copy of the rels part of the
		// layout of the source master (the relationship of the slide master
		// of the layout is rewritten to the new master)
		newLayoutRels := &relationships{}
		if sourceLayoutRels, _ := f.relsReader(f.getSlideRelsPath(oldLayoutPath)); sourceLayoutRels != nil {
			sourceLayoutRels.mu.Lock()
			for _, rel := range sourceLayoutRels.Relationships {
				if rel.Type == SourceRelationshipSlideMaster {
					rel.Target = "../slideMasters/" + newMasterBase
				}
				newLayoutRels.Relationships = append(newLayoutRels.Relationships, rel)
			}
			sourceLayoutRels.mu.Unlock()
		}
		f.Relationships.Store(f.getSlideRelsPath(newLayoutPath), newLayoutRels)
		_ = f.setContentTypes("/"+newLayoutPath, ContentTypeSlideLayoutML)
		_ = f.setContentTypes("/"+f.getSlideRelsPath(newLayoutPath), ContentTypeRelationships)

		copies = append(copies, layoutCopy{
			oldID:     oldID,
			newID:     nextLayoutID,
			oldTarget: oldTarget,
			newTarget: "../slideLayouts/" + filepath.Base(newLayoutPath),
		})
		nextLayoutID++
	}

	// the part of the new master: the byte copy of the part of the source
	// master (the ids of the slide layouts of the new master in the slide
	// layout id list are rewritten, the rIds of the slide layouts are kept)
	newMasterRaw := sourceMasterRaw
	for _, copy := range copies {
		re := regexp.MustCompile(`(<[A-Za-z0-9]+:sldLayoutId\s+id=")` + strconv.Itoa(copy.oldID) + `(")`)
		newMasterRaw = re.ReplaceAll(newMasterRaw, []byte("${1}"+strconv.Itoa(copy.newID)+"${2}"))
	}
	f.Pkg.Store(newMasterPath, newMasterRaw)

	// the theme of the source master: the new master gets its own copy of
	// the theme part (the theme part is private to the slide master in
	// PowerPoint: two slide masters sharing one theme part make the repair
	// prompt of PowerPoint appear)
	themeTarget := ""
	if sourceMasterRels != nil {
		sourceMasterRels.mu.Lock()
		for _, rel := range sourceMasterRels.Relationships {
			if rel.Type == SourceRelationshipTheme {
				themeTarget = rel.Target
				break
			}
		}
		sourceMasterRels.mu.Unlock()
	}
	newThemeBase := ""
	if themeTarget != "" {
		sourceThemePath := resolvePartPath(source.master, themeTarget)
		newThemePath := "ppt/theme/theme" + strconv.Itoa(f.nextThemePartIndex()) + ".xml"
		newThemeBase = filepath.Base(newThemePath)
		f.Pkg.Store(newThemePath, f.readBytes(sourceThemePath))
		_ = f.setContentTypes("/"+newThemePath, ContentTypeThemeML)
	}

	// the rels part of the new master: the copy of the rels part of the
	// source master (the targets of the relationships of the slide layouts
	// are rewritten to the parts of the new layouts, the target of the
	// relationship of the theme is rewritten to the part of the new theme,
	// the other relationships — the hyperlinks — are kept)
	newMasterRels := &relationships{}
	if sourceMasterRels != nil {
		sourceMasterRels.mu.Lock()
		for _, rel := range sourceMasterRels.Relationships {
			for _, copy := range copies {
				if rel.Target == copy.oldTarget {
					rel.Target = copy.newTarget
				}
			}
			if newThemeBase != "" && rel.Type == SourceRelationshipTheme {
				rel.Target = "../theme/" + newThemeBase
			}
			newMasterRels.Relationships = append(newMasterRels.Relationships, rel)
		}
		sourceMasterRels.mu.Unlock()
	}
	f.Relationships.Store(newMasterRelsPath, newMasterRels)
	_ = f.setContentTypes("/"+newMasterPath, ContentTypeSlideMasterML)
	_ = f.setContentTypes("/"+newMasterRelsPath, ContentTypeRelationships)

	// the slide master id list of the presentation: the id of the new master
	// (the relationship of the new master in the rels of the presentation)
	rID := f.addRels(f.getPresentationRelsPath(), SourceRelationshipSlideMaster,
		"slideMasters/"+newMasterBase, "")
	presentation, err := f.presentationReader()
	if err != nil {
		return -1, err
	}
	presentation.MasterSlide.MasterSlide = append(presentation.MasterSlide.MasterSlide, decodeSlideID{
		SlideID:        nextID,
		RelationshipID: "rId" + strconv.Itoa(rID),
	})

	return nextID, nil
}
