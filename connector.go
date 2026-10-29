// Copyright 2026 kenny-not-dead. All rights reserved.
// Use of this source code is governed by a BSD-style license
// that can be found in the LICENSE file.
//
// Package gopptx provides functionality to create and manipulate PowerPoint
// (.pptx) files in Go, using the Office Open XML (ECMA-376) format.

package gopptx

import "fmt"

// GetConnectors provides a function to get connectors (p:cxnSp shapes) by
// given slide id, in the document order. The connectors are returned by
// value: edits of the returned slice do not affect the slide.
func (f *File) GetConnectors(slideID int) ([]decodeConnector, error) {
	slide, err := f.slideReader(slideID)
	if err != nil {
		return nil, err
	}

	return slide.getConnectors(), nil
}

// getConnectors returns the connectors of the slide, in the document order.
func (ds *decodeSlide) getConnectors() []decodeConnector {
	var connectors []decodeConnector
	for i := range ds.CommonSlideData.ShapeTree.Children {
		if child := ds.CommonSlideData.ShapeTree.Children[i].Connector; child != nil {
			connectors = append(connectors, *child)
		}
	}

	return connectors
}

// GetConnector provides a function to get a connector by given slide id and
// connector id (the id attribute of the p:cNvPr element).
func (f *File) GetConnector(slideID int, connectorID int) (decodeConnector, error) {
	connectors, err := f.GetConnectors(slideID)
	if err != nil {
		return decodeConnector{}, err
	}

	for _, conn := range connectors {
		if conn.NonVisualConnectorProperties != nil &&
			conn.NonVisualConnectorProperties.CommonNonVisualProperties != nil &&
			conn.NonVisualConnectorProperties.CommonNonVisualProperties.ID == connectorID {
			return conn, nil
		}
	}

	return decodeConnector{}, ErrConnectorNotExist{connectorID}
}

// CreateConnector provides the function to create a new connector
// (p:cxnSp shape) by given slide id and returns the id of the connector in
// the slide after it was appended.
func (f *File) CreateConnector(slideID int, shapeProperties DecodeShapeProperties) (int, error) {
	slide, err := f.slideReader(slideID)
	if err != nil {
		return -1, err
	}

	// the connector id must be unique in the whole shape tree (sp + pic +
	// cxnSp), the same way CreateShape and AddImage allocate shape ids
	connectorID := f.nextShapeID(slide)

	name := fmt.Sprintf("Connector %d", connectorID)
	newConnector := decodeConnector{
		NonVisualConnectorProperties: &decodeNonVisualConnectorProperties{
			CommonNonVisualProperties: &decodeCommonNonVisualProperties{
				ID:   connectorID,
				Name: name,
			},
			// the connector non-visual properties and the non-visual
			// properties elements are expected by the consumers (the same
			// way CreateShape writes the empty elements of the shape)
			CommonNonVisualConnectorProperties: &CommonNonVisualConnectorProperties{},
			NonVisualProperties:                &decodeNonVisualProperties{},
		},
		ShapeProperties: &shapeProperties,
	}

	// the new connector is appended to the end of the shape tree: it is
	// drawn above the existing shapes of the slide
	slide.CommonSlideData.ShapeTree.Children = append(slide.CommonSlideData.ShapeTree.Children, decodeShapeTreeChild{Connector: &newConnector})

	return connectorID, nil
}

// DeleteConnector provides a function to delete a connector on the slide by
// given slide id and connector id (the id attribute of the p:cNvPr element).
func (f *File) DeleteConnector(slideID int, connectorID int) error {
	slide, err := f.slideReader(slideID)
	if err != nil {
		return err
	}

	children := slide.CommonSlideData.ShapeTree.Children

	deleteIndex := -1
	for i := range children {
		if conn := children[i].Connector; conn != nil &&
			conn.NonVisualConnectorProperties != nil &&
			conn.NonVisualConnectorProperties.CommonNonVisualProperties != nil &&
			conn.NonVisualConnectorProperties.CommonNonVisualProperties.ID == connectorID {

			deleteIndex = i

			break
		}
	}

	if deleteIndex == -1 {
		return ErrConnectorNotExist{connectorID}
	}

	slide.CommonSlideData.ShapeTree.Children = append(children[:deleteIndex], children[deleteIndex+1:]...)

	return nil
}
