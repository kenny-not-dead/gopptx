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
	"testing"

	"github.com/stretchr/testify/assert"
)

// TestGetConnectors checks the get of the connectors of the slide by the
// slide id: the slide without the connectors returns the empty list, the
// created connectors are returned in the document order, the slide that is
// not in the presentation returns the error.
func TestGetConnectors(t *testing.T) {
	f, err := OpenFile(testFixture)
	assert.NoError(t, err)
	defer f.Close()

	// the slide without the connectors
	connectors, err := f.GetConnectors(257)
	assert.NoError(t, err)
	assert.Empty(t, connectors)

	// the slide that is not in the presentation
	_, err = f.GetConnectors(999)
	assert.EqualError(t, err, ErrSlideNotExist{SlideID: 999}.Error())
}

// TestGetConnector checks the get of the connector by the slide id and the
// connector id: the created connector is returned with the name and the
// properties, the connector that is not in the slide returns the error.
func TestGetConnector(t *testing.T) {
	f, err := OpenFile(testFixture)
	assert.NoError(t, err)
	defer f.Close()

	// the connector that is not in the slide
	_, err = f.GetConnector(257, 999)
	assert.EqualError(t, err, ErrConnectorNotExist{ConnectorID: 999}.Error())

	shapeProperties := DecodeShapeProperties{
		Xfrm: &DecodeXfrm{
			Offset:  &Offset{X: 1000000, Y: 1000000},
			Extents: &Extents{CX: 3000000, CY: 0},
		},
	}
	connectorID, err := f.CreateConnector(257, shapeProperties)
	assert.NoError(t, err)

	connector, err := f.GetConnector(257, connectorID)
	assert.NoError(t, err)
	assert.Equal(t, connectorID, connector.NonVisualConnectorProperties.CommonNonVisualProperties.ID)
	assert.Equal(t, "Connector "+strconv.Itoa(connectorID), connector.NonVisualConnectorProperties.CommonNonVisualProperties.Name)
	assert.Equal(t, 1000000, connector.ShapeProperties.Xfrm.Offset.X)
	assert.Equal(t, 3000000, connector.ShapeProperties.Xfrm.Extents.CX)

	// the slide that is not in the presentation
	_, err = f.GetConnector(999, connectorID)
	assert.EqualError(t, err, ErrSlideNotExist{SlideID: 999}.Error())
}

// TestCreateConnector checks the create of the connector by the slide id:
// the id of the created connector is unique in the whole shape tree of the
// slide (the shapes, the pictures, the connectors and the groups), the
// created connectors are kept in the document order and the connector is
// preserved on the round trip.
func TestCreateConnector(t *testing.T) {
	f, err := OpenFile(testFixture)
	assert.NoError(t, err)
	defer f.Close()

	// the slide that is not in the presentation
	_, err = f.CreateConnector(999, DecodeShapeProperties{})
	assert.EqualError(t, err, ErrSlideNotExist{SlideID: 999}.Error())

	// the ids of the created connectors are the max id of the shape tree of
	// the slide plus one
	firstID, err := f.CreateConnector(257, DecodeShapeProperties{})
	assert.NoError(t, err)
	secondID, err := f.CreateConnector(257, DecodeShapeProperties{})
	assert.NoError(t, err)
	assert.Equal(t, firstID+1, secondID)

	// the connectors are returned in the document order
	connectors, err := f.GetConnectors(257)
	assert.NoError(t, err)
	assert.Len(t, connectors, 2)
	assert.Equal(t, firstID, connectors[0].NonVisualConnectorProperties.CommonNonVisualProperties.ID)
	assert.Equal(t, secondID, connectors[1].NonVisualConnectorProperties.CommonNonVisualProperties.ID)

	// the connector is preserved on the round trip
	out := filepath.Join(t.TempDir(), "connectors.pptx")
	assert.NoError(t, f.SaveAs(out))
	assert.Equal(t, 2, strings.Count(string(readZipPart(t, out, "ppt/slides/slide2.xml")), "<p:cxnSp>"))

	reopened, err := OpenFile(out)
	assert.NoError(t, err)
	defer reopened.Close()
	_, err = reopened.GetConnector(257, firstID)
	assert.NoError(t, err)
	_, err = reopened.GetConnector(257, secondID)
	assert.NoError(t, err)
}

// TestDeleteConnector checks the delete of the connector by the slide id and
// the connector id: the connector that is not in the slide returns the
// error, the delete of the created connectors keeps the slide unchanged on
// the round trip.
func TestDeleteConnector(t *testing.T) {
	f, err := OpenFile(testFixture)
	assert.NoError(t, err)
	defer f.Close()

	// the connector that is not in the slide
	err = f.DeleteConnector(257, 999)
	assert.EqualError(t, err, ErrConnectorNotExist{ConnectorID: 999}.Error())

	// the slide that is not in the presentation
	err = f.DeleteConnector(999, 999)
	assert.EqualError(t, err, ErrSlideNotExist{SlideID: 999}.Error())

	// the delete of the created connector keeps the slide without the
	// connector
	connectorID, err := f.CreateConnector(257, DecodeShapeProperties{})
	assert.NoError(t, err)
	err = f.DeleteConnector(257, connectorID)
	assert.NoError(t, err)

	// the connector is deleted
	_, err = f.GetConnector(257, connectorID)
	assert.EqualError(t, err, ErrConnectorNotExist{ConnectorID: connectorID}.Error())

	// the add and the delete of the connector keep the slide unchanged on
	// the round trip
	out := filepath.Join(t.TempDir(), "connectors.pptx")
	assert.NoError(t, f.SaveAs(out))
	assert.NoError(t, assertPartsEqual(t, testFixture, out))
}
