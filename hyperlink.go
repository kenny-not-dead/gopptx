// Copyright 2026 kenny-not-dead. All rights reserved.
// Use of this source code is governed by a BSD-style license
// that can be found in the LICENSE file.
//
// Package gopptx provides functionality to create and manipulate PowerPoint
// (.pptx) files in Go, using the Office Open XML (ECMA-376) format.

package gopptx

import (
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
	// LinkTypeLocation is the internal hyperlink to a part of the
	// presentation (the slides, for example). It is supported by
	// GetShapeHyperLinks as a filter only; SetShapeHyperLink returns
	// ErrInvalidLinkType for it.
	LinkTypeLocation LinkType = "Location"
	// LinkTypeNone removes the hyperlink from the shape (SetShapeHyperLink)
	// or filters the shapes without hyperlinks (GetShapeHyperLinks).
	LinkTypeNone LinkType = "None"
)

// SetShapeHyperLink provides a function to set the hyperlink of a shape by
// the given slide id and shape id. The linkType defines the type of the
// hyperlink: LinkTypeExternal for a URL (a hyperlink relationship with the
// TargetMode "External" is added to the slide rels part) and LinkTypeNone to
// remove the hyperlink from the shape. For a text shape the hyperlink is set
// to all the runs of the shape text (to the end paragraph run properties of
// the first paragraph when the shape text has no runs), so the whole shape
// text becomes the link; for a picture the hyperlink is set to the common
// non-visual properties of the picture, so the whole picture becomes the
// link. The hyperlink element (a:hlinkClick) of the run properties and of
// the common non-visual properties is preserved on the round trip, so the
// links of the loaded presentation are not lost on save.
//
//	tooltip := "gopptx on GitHub"
//	if err := f.SetShapeHyperLink(slideID, shapeID,
//	    "https://github.com/kenny-not-dead/gopptx",
//	    gopptx.LinkTypeExternal, gopptx.HyperlinkOpts{Tooltip: &tooltip}); err != nil {
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

	switch linkType {
	case LinkTypeExternal:
		hlink, err := f.shapeHyperlinkRelationship(slideID, link)
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
			pic.NonVisualPictureProperties.CommonNonVisualProperties.HlinkClick = &DecodeHyperlink{
				ID:      hlink.ID,
				Tooltip: hlink.Tooltip,
			}
		}
	case LinkTypeNone:
		if shape != nil {
			clearShapeRunHyperlinks(shape)
		}

		if pic != nil {
			pic.NonVisualPictureProperties.CommonNonVisualProperties.HlinkClick = nil
		}

		slidePath, _ := f.getSlideXMLPath(slideID)
		f.removeUnusedHyperlinkRels(slide, slidePath)
	default:
		return ErrInvalidLinkType
	}

	return nil
}

// GetShapeHyperLink gets the hyperlink of a shape by the given slide id and
// shape id. It returns true and the link address (the target of the
// hyperlink relationship) when the shape has a hyperlink, otherwise it
// returns false and an empty link address.
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
	rels, _ := f.relsReader(f.getSlideRelsPath(slidePath))
	if rels == nil {
		return false, "", nil
	}

	rels.mu.Lock()
	defer rels.mu.Unlock()
	for _, rel := range rels.Relationships {
		if rel.ID == id {
			return true, rel.Target, nil
		}
	}

	return false, "", nil
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
	rels, _ := f.relsReader(f.getSlideRelsPath(slidePath))
	isExternal := func(rID string) bool {
		if rID == "" || rels == nil {
			return false
		}

		rels.mu.Lock()
		defer rels.mu.Unlock()
		for _, rel := range rels.Relationships {
			if rel.ID == rID {
				return rel.Type == SourceRelationshipHyperlink
			}
		}

		return false
	}

	for i := range slide.CommonSlideData.ShapeTree.Children {
		child := &slide.CommonSlideData.ShapeTree.Children[i]
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
			if hasLink && isExternal(id) {
				ids = append(ids, shapeID)
			}
		case LinkTypeLocation:
			if hasLink && !isExternal(id) {
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

	return ids, nil
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
	// the hyperlink element references the relationship (r:id), so the slide
	// root element must declare the relationships namespace
	f.addNameSpaces(slidePath, SourceRelationship)
	relsPath := f.getSlideRelsPath(slidePath)
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

	return &Hyperlink{ID: "rId" + strconv.Itoa(rID)}, nil
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
		(*rp).HlinkClick = &DecodeHyperlink{ID: hlink.ID, Tooltip: hlink.Tooltip}
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
				rp.HlinkClick = nil
			}
		}
		if rp := paragraph.EndParagraphRunProperties; rp != nil {
			rp.HlinkClick = nil
		}
	}
}

// removeUnusedHyperlinkRels provides a function to remove the hyperlink
// relationships that are not used by any shape of the slide from the slide
// rels part.
func (f *File) removeUnusedHyperlinkRels(slide *decodeSlide, slidePath string) {
	used := map[string]bool{}
	mark := func(id string) {
		if id != "" {
			used[id] = true
		}
	}
	for i := range slide.CommonSlideData.ShapeTree.Children {
		child := &slide.CommonSlideData.ShapeTree.Children[i]
		markChild := func(cnv *decodeCommonNonVisualProperties) {
			if cnv == nil {
				return
			}
			if cnv.HlinkClick != nil {
				mark(cnv.HlinkClick.ID)
			}
			if cnv.HlinkMouseOver != nil {
				mark(cnv.HlinkMouseOver.ID)
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
							if rp.HlinkClick != nil {
								mark(rp.HlinkClick.ID)
							}
							if rp.HlinkMouseOver != nil {
								mark(rp.HlinkMouseOver.ID)
							}
						}
					}
					if rp := paragraph.EndParagraphRunProperties; rp != nil {
						if rp.HlinkClick != nil {
							mark(rp.HlinkClick.ID)
						}
						if rp.HlinkMouseOver != nil {
							mark(rp.HlinkMouseOver.ID)
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

	rels, _ := f.relsReader(f.getSlideRelsPath(slidePath))
	if rels == nil {
		return
	}
	rels.mu.Lock()
	defer rels.mu.Unlock()
	for idx := len(rels.Relationships) - 1; idx >= 0; idx-- {
		rel := &rels.Relationships[idx]
		if rel.Type == SourceRelationshipHyperlink && !used[rel.ID] {
			rels.Relationships = append(rels.Relationships[:idx], rels.Relationships[idx+1:]...)
		}
	}
}

// shapeHyperlinkID provides a function to get the id of the hyperlink
// relationship of the shape: first the common non-visual properties (the
// click-through link), then the runs of the shape text. When the runs
// contain different hyperlinks, it returns ErrShapeHyperlinkAmbiguous.
func shapeHyperlinkID(shape *decodeShape, pic *decodePicture) (string, error) {
	if pic != nil && pic.NonVisualPictureProperties != nil &&
		pic.NonVisualPictureProperties.CommonNonVisualProperties != nil {
		if hlink := pic.NonVisualPictureProperties.CommonNonVisualProperties.HlinkClick; hlink != nil {
			return hlink.ID, nil
		}
	}

	if shape == nil {
		return "", nil
	}

	if shape.NonVisualShapeProperties != nil &&
		shape.NonVisualShapeProperties.CommonNonVisualProperties != nil {
		if hlink := shape.NonVisualShapeProperties.CommonNonVisualProperties.HlinkClick; hlink != nil {
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
			if rp := paragraph.Runs[r].RunProperties; rp != nil && rp.HlinkClick != nil {
				if id == "" {
					id = rp.HlinkClick.ID
				} else if id != rp.HlinkClick.ID {
					return "", ErrShapeHyperlinkAmbiguous
				}
			}
		}
		if rp := paragraph.EndParagraphRunProperties; rp != nil && rp.HlinkClick != nil {
			if id == "" {
				id = rp.HlinkClick.ID
			} else if id != rp.HlinkClick.ID {
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
