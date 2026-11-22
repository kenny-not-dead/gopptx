// Copyright 2026 kenny-not-dead. All rights reserved.
// Use of this source code is governed by a BSD-style license
// that can be found in the LICENSE file.
//
// Package gopptx provides functionality to create and manipulate PowerPoint
// (.pptx) files in Go, using the Office Open XML (ECMA-376) format.

package gopptx

import (
	"path/filepath"
	"strconv"
	"strings"
)

// HyperlinkOpts can be passed to SetShapeHyperLink to set the optional
// attributes of the hyperlink (the tooltip shown on hover).
type HyperlinkOpts struct {
	Tooltip *string
}

// LinkType defines the type of the shape hyperlink.
type LinkType string

const (
	// LinkTypeExternal is the URL hyperlink: a hyperlink relationship with
	// the TargetMode "External" is added to the slide rels part of the
	// slide.
	LinkTypeExternal LinkType = "External"
	// LinkTypeLocation is the internal hyperlink to the slide of the
	// presentation: the link is the id of the target slide, a relationship
	// of the type "slide" (without the TargetMode) is added to the slide
	// rels part of the slide and the "ppaction://hlinksldjump" action is
	// written to the hyperlink element, so the link jumps to the target
	// slide in the presentation.
	LinkTypeLocation LinkType = "Location"
	// LinkTypeNone removes the hyperlink from the shape (SetShapeHyperLink)
	// or filters the shapes without hyperlinks (GetShapeHyperLinks).
	LinkTypeNone LinkType = "None"
)

// SetShapeHyperLink provides a function to set the hyperlink of a shape by
// the given slide id and shape id. The linkType defines the type of the
// hyperlink: LinkTypeExternal for a URL (a hyperlink relationship with the
// TargetMode "External" is added to the slide rels part), LinkTypeLocation
// for the internal link to the slide of the presentation (the link is the id
// of the target slide, a relationship of the type "slide" is added to the
// slide rels part and the "ppaction://hlinksldjump" action is written to the
// hyperlink element, so the link jumps to the target slide in the
// presentation) and LinkTypeNone to remove the hyperlink from the shape. For
// a text shape the hyperlink is set to all the runs of the shape text (to
// the end paragraph run properties of the first paragraph when the shape
// text has no runs), so the whole shape text becomes the link; for a picture
// the hyperlink is set to the common non-visual properties of the picture,
// so the whole picture becomes the link. The hyperlink element (a:hlinkClick)
// of the run properties and of the common non-visual properties is preserved
// on the round trip, so the links of the loaded presentation are not lost on
// save.
//
//	tooltip := "gopptx on GitHub"
//	if err := f.SetShapeHyperLink(slideID, shapeID,
//	    "https://github.com/kenny-not-dead/gopptx",
//	    gopptx.LinkTypeExternal, gopptx.HyperlinkOpts{Tooltip: &tooltip}); err != nil {
//	    fmt.Println(err)
//	}
//
//	the internal link to the slide with the id 257:
//
//	if err := f.SetShapeHyperLink(slideID, shapeID, "257", gopptx.LinkTypeLocation); err != nil {
//	    fmt.Println(err)
//	}
func (f *File) SetShapeHyperLink(slideID, shapeID int, link string, linkType LinkType, opts ...HyperlinkOpts) error {
	slide, err := f.slideReader(slideID)
	if err != nil {
		return err
	}

	var (
		shape *decodeShape
		pic   *decodePicture
	)
	for i := range slide.CommonSlideData.ShapeTree.Children {
		child := &slide.CommonSlideData.ShapeTree.Children[i]
		if shapeIDMatch(child.Shape, shapeID) {
			shape = child.Shape
			break
		}
		if shapeIDMatchPicture(child.Picture, shapeID) {
			pic = child.Picture
			break
		}
	}

	if shape == nil && pic == nil {
		return ErrShapeNotExist{shapeID}
	}

	var hlink *Hyperlink
	switch linkType {
	case LinkTypeExternal:
		hlink, err = f.shapeHyperlinkRelationship(slideID, link)
	case LinkTypeLocation:
		hlink, err = f.shapeLocationRelationship(slideID, link)
	case LinkTypeNone:
		if shape != nil {
			clearShapeRunHyperlinks(shape)
		}

		if pic != nil {
			pic.NonVisualPictureProperties.CommonNonVisualProperties.HyperlinkClick = nil
		}

		slidePath, _ := f.getSlideXMLPath(slideID)
		f.removeUnusedLinkRels(slidePath, shapeTreeUsedLinkIDs(slide.CommonSlideData.ShapeTree.Children))

		return nil
	default:
		return ErrInvalidLinkType
	}

	if err != nil {
		return err
	}

	for _, o := range opts {
		if o.Tooltip != nil {
			hlink.Tooltip = *o.Tooltip
		}
	}
	if shape != nil {
		setShapeRunHyperlinks(shape, hlink)
	} else {
		pic.NonVisualPictureProperties.CommonNonVisualProperties.HyperlinkClick = &DecodeHyperlink{
			ID:      hlink.ID,
			Action:  hlink.Action,
			Tooltip: hlink.Tooltip,
		}
	}

	return nil
}

// GetShapeHyperLink gets the hyperlink of a shape by the given slide id and
// shape id. It returns true and the link address when the shape has a
// hyperlink, otherwise it returns false and an empty link address: the
// target of the hyperlink relationship for the URL links (LinkTypeExternal)
// and the id of the target slide for the internal links (LinkTypeLocation).
//
// For example, get the hyperlink of the shape with the id 7 on the slide
// with the id 256:
//
//	hasLink, link, err := f.GetShapeHyperLink(256, 7)
func (f *File) GetShapeHyperLink(slideID, shapeID int) (bool, string, error) {
	slide, err := f.slideReader(slideID)
	if err != nil {
		return false, "", err
	}

	var (
		shape *decodeShape
		pic   *decodePicture
	)
	for i := range slide.CommonSlideData.ShapeTree.Children {
		child := &slide.CommonSlideData.ShapeTree.Children[i]
		if shapeIDMatch(child.Shape, shapeID) {
			shape = child.Shape
			break
		}

		if shapeIDMatchPicture(child.Picture, shapeID) {
			pic = child.Picture
			break
		}
	}

	if shape == nil && pic == nil {
		return false, "", ErrShapeNotExist{shapeID}
	}

	id, err := shapeHyperlinkID(shape, pic)
	if err != nil {
		return false, "", err
	}

	if id == "" {
		return false, "", nil
	}

	slidePath, _ := f.getSlideXMLPath(slideID)
	return f.partHyperlinkTarget(slidePath, f.getSlideRelsPath(slidePath), id)
}

// GetShapeHyperLinks returns the shape ids that contain hyperlinks in the
// given slide. The optional parameter linkType is used to filter the links:
// LinkTypeExternal for the URL links, LinkTypeLocation for the internal
// links to the parts of the presentation and LinkTypeNone for the shapes
// without hyperlinks. When linkType is empty, all the shapes with
// hyperlinks are returned.
func (f *File) GetShapeHyperLinks(slideID int, linkType LinkType) ([]int, error) {
	var ids []int
	slide, err := f.slideReader(slideID)
	if err != nil {
		return ids, err
	}

	switch linkType {
	case "", LinkTypeExternal, LinkTypeLocation, LinkTypeNone:
	default:
		return ids, ErrInvalidLinkType
	}

	slidePath, _ := f.getSlideXMLPath(slideID)
	return f.shapeTreeHyperlinkIDs(slide.CommonSlideData.ShapeTree.Children, f.getSlideRelsPath(slidePath), linkType), nil
}

// shapeIDMatch reports whether the given shape or picture (or both are nil)
// has the given shape id.
func shapeIDMatch(shape *decodeShape, shapeID int) bool {
	if shape != nil && shape.NonVisualShapeProperties != nil &&
		shape.NonVisualShapeProperties.CommonNonVisualProperties != nil {
		return shape.NonVisualShapeProperties.CommonNonVisualProperties.ID == shapeID
	}

	return false
}

func shapeIDMatchPicture(pic *decodePicture, shapeID int) bool {
	if pic != nil && pic.NonVisualPictureProperties != nil &&
		pic.NonVisualPictureProperties.CommonNonVisualProperties != nil {
		return pic.NonVisualPictureProperties.CommonNonVisualProperties.ID == shapeID
	}

	return false
}

// shapeHyperlinkRelationship provides a function to add the hyperlink
// relationship to the slide rels part and returns the hyperlink element with
// the relationship id. An existing hyperlink relationship with the same
// target is reused.
func (f *File) shapeHyperlinkRelationship(slideID int, target string) (*Hyperlink, error) {
	slidePath, _ := f.getSlideXMLPath(slideID)

	return f.partHyperlinkRelationship(slidePath, f.getSlideRelsPath(slidePath), target), nil
}

// partHyperlinkRelationship provides a function to add the hyperlink
// relationship to the rels part of the given part (the slide, the slide
// master or the slide layout) and returns the hyperlink element with the
// relationship id. An existing hyperlink relationship with the same target
// is reused.
func (f *File) partHyperlinkRelationship(partPath, relsPath, target string) *Hyperlink {
	// the hyperlink element references the relationship (r:id), so the part
	// root element must declare the relationships namespace
	f.addNameSpaces(partPath, SourceRelationship)
	rels, _ := f.relsReader(relsPath)
	rID := 0
	if rels != nil {
		rels.mu.Lock()
		for _, rel := range rels.Relationships {
			if rel.Type == SourceRelationshipHyperlink && rel.Target == target {
				rID, _ = strconv.Atoi(strings.TrimPrefix(rel.ID, "rId"))
				break
			}
		}
		rels.mu.Unlock()
	}

	if rID == 0 {
		rID = f.addRels(relsPath, SourceRelationshipHyperlink, target, "External")
	}

	return &Hyperlink{ID: "rId" + strconv.Itoa(rID)}
}

// actionSlideJump is the action of the internal hyperlink to the slide of
// the presentation (the "ppaction://hlinksldjump" action of the
// a:hlinkClick element) that PowerPoint writes for the links to the slides
// of the presentation.
const actionSlideJump = "ppaction://hlinksldjump"

// shapeLocationRelationship provides a function to add the relationship to
// the target slide of the internal hyperlink to the slide rels part and
// returns the hyperlink element with the relationship id and the slide jump
// action. The link is the id of the target slide. An existing relationship
// with the same target is reused.
func (f *File) shapeLocationRelationship(slideID int, link string) (*Hyperlink, error) {
	targetSlideID, err := strconv.Atoi(link)
	if err != nil {
		return nil, ErrLinkTargetInvalid{Target: link}
	}
	slidePath, _ := f.getSlideXMLPath(slideID)

	return f.partLocationRelationship(slidePath, f.getSlideRelsPath(slidePath), targetSlideID)
}

// partLocationRelationship provides a function to add the relationship to
// the target slide of the internal hyperlink to the rels part of the given
// part (the slide, the slide master or the slide layout) and returns the
// hyperlink element with the relationship id and the slide jump action. The
// target of the relationship is the relative path of the target slide part
// (the part name for the slide, the "../slides/slideN.xml" path for the
// slide master and the slide layout). An existing relationship with the same
// target is reused.
func (f *File) partLocationRelationship(partPath, relsPath string, targetSlideID int) (*Hyperlink, error) {
	targetPath, ok := f.getSlideXMLPath(targetSlideID)
	if !ok {
		return nil, ErrSlideNotExist{SlideID: targetSlideID}
	}

	// the hyperlink element references the relationship (r:id), so the part
	// root element must declare the relationships namespace
	f.addNameSpaces(partPath, SourceRelationship)
	// the target of the relationship is the path of the target slide part,
	// relative to the directory of the part
	target, err := filepath.Rel(filepath.ToSlash(filepath.Dir(partPath)), targetPath)
	if err != nil {
		return nil, err
	}
	rels, _ := f.relsReader(relsPath)
	rID := 0
	if rels != nil {
		rels.mu.Lock()
		for _, rel := range rels.Relationships {
			if rel.Type == SourceRelationshipSlide && rel.Target == target {
				rID, _ = strconv.Atoi(strings.TrimPrefix(rel.ID, "rId"))
				break
			}
		}
		rels.mu.Unlock()
	}

	if rID == 0 {
		rID = f.addRels(relsPath, SourceRelationshipSlide, target, "")
	}

	return &Hyperlink{ID: "rId" + strconv.Itoa(rID), Action: actionSlideJump}, nil
}

// setShapeRunHyperlinks provides a function to set the hyperlink element to
// all the runs of the shape text (to the end paragraph run properties of the
// first paragraph when the shape text has no runs), the same as PowerPoint
// does for the linked text.
func setShapeRunHyperlinks(shape *decodeShape, hlink *Hyperlink) {
	if shape.TextBody == nil {
		return
	}

	set := func(rp **DecodeRunProperties) {
		if *rp == nil {
			*rp = new(DecodeRunProperties)
		}
		(*rp).HyperlinkClick = &DecodeHyperlink{ID: hlink.ID, Action: hlink.Action, Tooltip: hlink.Tooltip}
	}
	runCount := 0
	for p := range shape.TextBody.Paragraph {
		paragraph := &shape.TextBody.Paragraph[p]
		for r := range paragraph.Runs {
			set(&paragraph.Runs[r].RunProperties)
		}
		runCount += len(paragraph.Runs)
	}
	if runCount == 0 && len(shape.TextBody.Paragraph) > 0 {
		set(&shape.TextBody.Paragraph[0].EndParagraphRunProperties)
	}
}

// clearShapeRunHyperlinks provides a function to remove the hyperlink element
// from all the runs of the shape text and from the end paragraph run
// properties of the paragraphs.
func clearShapeRunHyperlinks(shape *decodeShape) {
	if shape.TextBody == nil {
		return
	}
	for p := range shape.TextBody.Paragraph {
		paragraph := &shape.TextBody.Paragraph[p]
		for r := range paragraph.Runs {
			if rp := paragraph.Runs[r].RunProperties; rp != nil {
				rp.HyperlinkClick = nil
			}
		}
		if rp := paragraph.EndParagraphRunProperties; rp != nil {
			rp.HyperlinkClick = nil
		}
	}
}

// shapeTreeUsedLinkIDs provides a function to get the ids of the hyperlink
// relationships (the URLs links) and the relationships to the slides (the
// internal links) that are used by the shapes, the pictures and the
// connectors of the shape tree.
func shapeTreeUsedLinkIDs(children []decodeShapeTreeChild) map[string]bool {
	used := map[string]bool{}
	mark := func(id string) {
		if id != "" {
			used[id] = true
		}
	}
	for i := range children {
		child := &children[i]
		markChild := func(cnv *decodeCommonNonVisualProperties) {
			if cnv == nil {
				return
			}
			if cnv.HyperlinkClick != nil {
				mark(cnv.HyperlinkClick.ID)
			}
			if cnv.HyperlinkMouseOver != nil {
				mark(cnv.HyperlinkMouseOver.ID)
			}
		}
		if child.Shape != nil {
			if child.Shape.NonVisualShapeProperties != nil {
				markChild(child.Shape.NonVisualShapeProperties.CommonNonVisualProperties)
			}
			if child.Shape.TextBody != nil {
				for p := range child.Shape.TextBody.Paragraph {
					paragraph := &child.Shape.TextBody.Paragraph[p]
					for r := range paragraph.Runs {
						if rp := paragraph.Runs[r].RunProperties; rp != nil {
							if rp.HyperlinkClick != nil {
								mark(rp.HyperlinkClick.ID)
							}
							if rp.HyperlinkMouseOver != nil {
								mark(rp.HyperlinkMouseOver.ID)
							}
						}
					}
					if rp := paragraph.EndParagraphRunProperties; rp != nil {
						if rp.HyperlinkClick != nil {
							mark(rp.HyperlinkClick.ID)
						}
						if rp.HyperlinkMouseOver != nil {
							mark(rp.HyperlinkMouseOver.ID)
						}
					}
				}
			}
		}
		if child.Picture != nil && child.Picture.NonVisualPictureProperties != nil {
			markChild(child.Picture.NonVisualPictureProperties.CommonNonVisualProperties)
		}
		if child.Connector != nil && child.Connector.NonVisualConnectorProperties != nil {
			markChild(child.Connector.NonVisualConnectorProperties.CommonNonVisualProperties)
		}
	}

	return used
}

// removeUnusedLinkRels provides a function to remove the hyperlink
// relationships (the URL links) and the relationships to the slides (the
// internal links) that are not used by any shape of the part from the rels
// part of the given part (the slide, the slide master or the slide layout).
func (f *File) removeUnusedLinkRels(partPath string, used map[string]bool) {
	rels, _ := f.relsReader(f.getSlideRelsPath(partPath))
	if rels == nil {
		return
	}
	rels.mu.Lock()
	defer rels.mu.Unlock()
	for idx := len(rels.Relationships) - 1; idx >= 0; idx-- {
		rel := &rels.Relationships[idx]
		if (rel.Type == SourceRelationshipHyperlink || rel.Type == SourceRelationshipSlide) &&
			!used[rel.ID] {
			rels.Relationships = append(rels.Relationships[:idx], rels.Relationships[idx+1:]...)
		}
	}
}

// partHyperlinkTarget provides a function to get the target of the shape
// hyperlink by the relationship id in the rels part of the given part (the
// slide, the slide master or the slide layout): the target of the hyperlink
// relationship for the URL links (LinkTypeExternal) and the id of the target
// slide for the internal links (LinkTypeLocation). It returns false and an
// empty target when the relationship is not found.
func (f *File) partHyperlinkTarget(partPath, relsPath, rID string) (bool, string, error) {
	rels, _ := f.relsReader(relsPath)
	if rels == nil {
		return false, "", nil
	}
	rels.mu.Lock()
	defer rels.mu.Unlock()
	for _, rel := range rels.Relationships {
		if rel.ID == rID {
			if rel.Type == SourceRelationshipSlide {
				// the internal link: the address is the id of the target slide
				if target := f.slideIDByPart(resolvePartPath(partPath, rel.Target)); target != 0 {
					return true, strconv.Itoa(target), nil
				}
			}

			return true, rel.Target, nil
		}
	}

	return false, "", nil
}

// shapeTreeHyperlinkIDs provides a function to get the shape ids that contain
// hyperlinks in the shape tree (the slide, the slide master or the slide
// layout part). The linkType filters the links: LinkTypeExternal for the URL
// links, LinkTypeLocation for the internal links to the parts of the
// presentation and LinkTypeNone for the shapes without hyperlinks. When the
// linkType is empty, all the shapes with hyperlinks are returned.
func (f *File) shapeTreeHyperlinkIDs(children []decodeShapeTreeChild, relsPath string, linkType LinkType) []int {
	var ids []int
	rels, _ := f.relsReader(relsPath)
	// relType returns the type of the relationship of the hyperlink of the
	// shape (the type "hyperlink" for the URL links and the type "slide" for
	// the internal links) or an empty type when the shape has no hyperlink
	relType := func(rID string) string {
		if rID == "" || rels == nil {
			return ""
		}

		rels.mu.Lock()
		defer rels.mu.Unlock()
		for _, rel := range rels.Relationships {
			if rel.ID == rID {
				return rel.Type
			}
		}

		return ""
	}

	for i := range children {
		child := &children[i]
		shapeID := 0
		if child.Shape != nil && child.Shape.NonVisualShapeProperties != nil &&
			child.Shape.NonVisualShapeProperties.CommonNonVisualProperties != nil {
			shapeID = child.Shape.NonVisualShapeProperties.CommonNonVisualProperties.ID
		}
		if child.Picture != nil && child.Picture.NonVisualPictureProperties != nil &&
			child.Picture.NonVisualPictureProperties.CommonNonVisualProperties != nil {
			shapeID = child.Picture.NonVisualPictureProperties.CommonNonVisualProperties.ID
		}
		id := childHyperlinkID(child)
		hasLink := id != ""
		switch linkType {
		case LinkTypeExternal:
			if hasLink && relType(id) == SourceRelationshipHyperlink {
				ids = append(ids, shapeID)
			}
		case LinkTypeLocation:
			if hasLink && relType(id) == SourceRelationshipSlide {
				ids = append(ids, shapeID)
			}
		case LinkTypeNone:
			if !hasLink {
				ids = append(ids, shapeID)
			}
		default:
			if hasLink {
				ids = append(ids, shapeID)
			}
		}
	}

	return ids
}

// shapeHyperlinkID provides a function to get the id of the hyperlink
// relationship of the shape: first the common non-visual properties (the
// click-through link), then the runs of the shape text. When the runs
// contain different hyperlinks, it returns ErrShapeHyperlinkAmbiguous.
func shapeHyperlinkID(shape *decodeShape, pic *decodePicture) (string, error) {
	if pic != nil && pic.NonVisualPictureProperties != nil &&
		pic.NonVisualPictureProperties.CommonNonVisualProperties != nil {
		if hlink := pic.NonVisualPictureProperties.CommonNonVisualProperties.HyperlinkClick; hlink != nil {
			return hlink.ID, nil
		}
	}

	if shape == nil {
		return "", nil
	}

	if shape.NonVisualShapeProperties != nil &&
		shape.NonVisualShapeProperties.CommonNonVisualProperties != nil {
		if hlink := shape.NonVisualShapeProperties.CommonNonVisualProperties.HyperlinkClick; hlink != nil {
			return hlink.ID, nil
		}
	}

	if shape.TextBody == nil {
		return "", nil
	}

	id := ""
	for p := range shape.TextBody.Paragraph {
		paragraph := &shape.TextBody.Paragraph[p]
		for r := range paragraph.Runs {
			if rp := paragraph.Runs[r].RunProperties; rp != nil && rp.HyperlinkClick != nil {
				if id == "" {
					id = rp.HyperlinkClick.ID
				} else if id != rp.HyperlinkClick.ID {
					return "", ErrShapeHyperlinkAmbiguous
				}
			}
		}
		if rp := paragraph.EndParagraphRunProperties; rp != nil && rp.HyperlinkClick != nil {
			if id == "" {
				id = rp.HyperlinkClick.ID
			} else if id != rp.HyperlinkClick.ID {
				return "", ErrShapeHyperlinkAmbiguous
			}
		}
	}

	return id, nil
}

// childHyperlinkID provides a function to get the id of the hyperlink
// relationship of a shape tree child (the shape or the picture), or an empty
// id when the child has no hyperlink.
func childHyperlinkID(child *decodeShapeTreeChild) string {
	if child.Picture != nil {
		id, _ := shapeHyperlinkID(nil, child.Picture)
		return id
	}

	if child.Shape != nil {
		id, _ := shapeHyperlinkID(child.Shape, nil)
		return id
	}

	return ""
}
