// Copyright 2026 kenny-not-dead. All rights reserved.
// Use of this source code is governed by a BSD-style license
// that can be found in the LICENSE file.
//
// Package gopptx provides functionality to create and manipulate PowerPoint
// (.pptx) files in Go, using the Office Open XML (ECMA-376) format.

package gopptx

import (
	"bytes"
	"strings"
)

// The shape tree helpers below operate on the common slide data (the p:cSld
// element, the CT_CommonSlideData of the ECMA-376 schema): the background and
// the shape tree that a slide, a slide layout and a slide master part all
// share. Because the three parts embed the same decodeSlideData, the
// shape-tree operations (find a shape by its id, list the shapes, delete a
// shape, the next free id) are the same for all of them, so they are defined
// once here, on the decodeSlideData type.

// findShape provides a function to find the shape by the given id in the shape
// tree of the common slide data (the slide, the slide layout or the slide
// master part). It returns nil when the shape is not found.
func (cs *decodeSlideData) findShape(shapeID int) *decodeShape {
	for i := range cs.ShapeTree.Children {
		child := &cs.ShapeTree.Children[i]
		if child.Shape == nil {
			continue
		}
		nv := child.Shape.NonVisualShapeProperties
		if nv == nil || nv.CommonNonVisualProperties == nil {
			continue
		}
		if nv.CommonNonVisualProperties.ID == shapeID {
			return child.Shape
		}
	}

	return nil
}

// shapes returns the shapes of the common slide data, in the document order.
func (cs *decodeSlideData) shapes() []decodeShape {
	var shapes []decodeShape
	for i := range cs.ShapeTree.Children {
		if child := cs.ShapeTree.Children[i].Shape; child != nil {
			shapes = append(shapes, *child)
		}
	}

	return shapes
}

// deleteShape provides a function to delete the shape by the given id from the
// shape tree of the common slide data. It returns false when the shape is not
// found.
func (cs *decodeSlideData) deleteShape(shapeID int) bool {
	for i := range cs.ShapeTree.Children {
		child := &cs.ShapeTree.Children[i]
		if child.Shape == nil {
			continue
		}
		nv := child.Shape.NonVisualShapeProperties
		if nv == nil || nv.CommonNonVisualProperties == nil {
			continue
		}
		if nv.CommonNonVisualProperties.ID == shapeID {
			children := cs.ShapeTree.Children
			cs.ShapeTree.Children = append(children[:i], children[i+1:]...)
			return true
		}
	}

	return false
}

// nextShapeID provides a function to get the next free id of the shape of the
// shape tree of the common slide data: the max id of the children of the shape
// tree (the shapes, the pictures and the connectors) plus one.
func (cs *decodeSlideData) nextShapeID() int {
	maxID := 0
	for i := range cs.ShapeTree.Children {
		child := &cs.ShapeTree.Children[i]
		var id int
		switch {
		case child.Shape != nil &&
			child.Shape.NonVisualShapeProperties != nil &&
			child.Shape.NonVisualShapeProperties.CommonNonVisualProperties != nil:
			id = child.Shape.NonVisualShapeProperties.CommonNonVisualProperties.ID
		case child.Picture != nil &&
			child.Picture.NonVisualPictureProperties != nil &&
			child.Picture.NonVisualPictureProperties.CommonNonVisualProperties != nil:
			id = child.Picture.NonVisualPictureProperties.CommonNonVisualProperties.ID
		case child.Connector != nil &&
			child.Connector.NonVisualConnectorProperties != nil &&
			child.Connector.NonVisualConnectorProperties.CommonNonVisualProperties != nil:
			id = child.Connector.NonVisualConnectorProperties.CommonNonVisualProperties.ID
		}
		if id > maxID {
			maxID = id
		}
	}

	return maxID + 1
}

// commonSlideDataUnsupportedElement checks the raw bytes of the common slide
// data of the slide layout or the slide master part for the elements that are
// not modeled by the layout and the master round trip (the shape tree must
// contain the p:sp shapes only, the markup compatibility element is not
// allowed in the common slide data). It returns the name of the unsupported
// element.
func commonSlideDataUnsupportedElement(raw []byte) string {
	for _, element := range []string{
		":grpSp>", ":custDataLst", ":controls", ":overrideClrMapping",
	} {
		if bytes.Contains(raw, []byte(element)) {
			return strings.TrimPrefix(element, ":")
		}
	}
	// the markup compatibility element of the common slide data is not
	// modeled by the layout and the master round trip
	if i := bytes.Index(raw, []byte("<p:cSld")); i != -1 {
		if j := bytes.Index(raw, []byte("</p:cSld>")); j > i {
			if bytes.Contains(raw[i:j], []byte("<mc:AlternateContent")) {
				return "the markup compatibility element of the common slide data"
			}
		}
	}
	return ""
}
