// Copyright 2026 kenny-not-dead. All rights reserved.
// Use of this source code is governed by a BSD-style license
// that can be found in the LICENSE file.
//
// Package gopptx provides functionality to create and manipulate PowerPoint
// (.pptx) files in Go, using the Office Open XML (ECMA-376) format.

package gopptx

import (
	"path/filepath"
	"strings"
	"testing"

	"github.com/stretchr/testify/assert"
)

// TestNewSlide checks the create of the new slide: the id of the new slide
// is the max id of the slides of the presentation plus one, the new slide is
// in the slide list and the slide part is created.
func TestNewSlide(t *testing.T) {
	f, err := OpenFile(testFixture)
	assert.NoError(t, err)
	defer f.Close()

	assert.Equal(t, []int{256, 257, 258}, f.GetSlideList())

	slideID, err := f.NewSlide()
	assert.NoError(t, err)
	assert.Equal(t, 259, slideID)
	assert.Equal(t, []int{256, 257, 258, 259}, f.GetSlideList())

	// the new slide part is created
	_, err = f.GetShapes(259)
	assert.NoError(t, err)

	out := filepath.Join(t.TempDir(), "new_slide.pptx")
	assert.NoError(t, f.SaveAs(out))

	reopened, err := OpenFile(out)
	assert.NoError(t, err)
	defer reopened.Close()
	assert.Equal(t, []int{256, 257, 258, 259}, reopened.GetSlideList())
}

// TestSetSlideBackground checks the set, the get and the delete of the slide
// background by the slide id: the background is replaced, the deleted
// background is inherited from the slide layout or the slide master, the
// background is preserved on the round trip, the slide that is not in the
// presentation returns the error.
func TestSetSlideBackground(t *testing.T) {
	f, err := OpenFile(testFixture)
	assert.NoError(t, err)
	defer f.Close()

	// the slide without the explicit background
	background, err := f.GetSlideBackground(257)
	assert.NoError(t, err)
	assert.Nil(t, background)

	// the slide that is not in the presentation
	err = f.SetSlideBackground(999, DecodeBackground{})
	assert.EqualError(t, err, ErrSlideNotExist{SlideID: 999}.Error())
	_, err = f.GetSlideBackground(999)
	assert.EqualError(t, err, ErrSlideNotExist{SlideID: 999}.Error())
	err = f.DeleteSlideBackground(999)
	assert.EqualError(t, err, ErrSlideNotExist{SlideID: 999}.Error())

	// the background is set to the slide
	err = f.SetSlideBackground(257, DecodeBackground{
		BGProperties: &DecodeBGProperties{
			SolidFill: &DecodeSolidFill{SolidRGBColor: &SolidRGBColor{Val: "FF0000"}},
		},
	})
	assert.NoError(t, err)
	background, err = f.GetSlideBackground(257)
	assert.NoError(t, err)
	assert.NotNil(t, background)
	assert.Equal(t, "FF0000", background.BGProperties.SolidFill.SolidRGBColor.Val)

	// the background is preserved on the round trip
	out := filepath.Join(t.TempDir(), "background.pptx")
	assert.NoError(t, f.SaveAs(out))
	reopened, err := OpenFile(out)
	assert.NoError(t, err)
	defer reopened.Close()
	background, err = reopened.GetSlideBackground(257)
	assert.NoError(t, err)
	assert.NotNil(t, background)
	assert.Equal(t, "FF0000", background.BGProperties.SolidFill.SolidRGBColor.Val)

	// the deleted background is inherited from the slide layout or the
	// slide master
	err = reopened.DeleteSlideBackground(257)
	assert.NoError(t, err)
	background, err = reopened.GetSlideBackground(257)
	assert.NoError(t, err)
	assert.Nil(t, background)
}

// TestSetShapeFill checks the set, the get and the delete of the shape fill
// by the slide id and the shape id: the fill is replaced, the deleted fill
// is replaced by the noFill property, the fill is preserved on the round
// trip, the shape that is not in the slide returns the error.
func TestSetShapeFill(t *testing.T) {
	f, err := OpenFile(testFixture)
	assert.NoError(t, err)
	defer f.Close()

	// the shape without the solid fill
	fill, err := f.GetShapeFill(257, 8)
	assert.NoError(t, err)
	assert.Nil(t, fill)

	// the shape that is not in the slide
	err = f.SetShapeFill(257, 999, DecodeSolidFill{})
	assert.EqualError(t, err, ErrShapeNotExist{ShapeID: 999}.Error())
	_, err = f.GetShapeFill(257, 999)
	assert.EqualError(t, err, ErrShapeNotExist{ShapeID: 999}.Error())
	err = f.DeleteShapeFill(257, 999)
	assert.EqualError(t, err, ErrShapeNotExist{ShapeID: 999}.Error())
	// the slide that is not in the presentation
	err = f.SetShapeFill(999, 8, DecodeSolidFill{})
	assert.EqualError(t, err, ErrSlideNotExist{SlideID: 999}.Error())
	_, err = f.GetShapeFill(999, 8)
	assert.EqualError(t, err, ErrSlideNotExist{SlideID: 999}.Error())
	err = f.DeleteShapeFill(999, 8)
	assert.EqualError(t, err, ErrSlideNotExist{SlideID: 999}.Error())

	// the fill is set to the shape
	err = f.SetShapeFill(257, 8, DecodeSolidFill{SolidRGBColor: &SolidRGBColor{Val: "00FF00"}})
	assert.NoError(t, err)
	fill, err = f.GetShapeFill(257, 8)
	assert.NoError(t, err)
	assert.NotNil(t, fill)
	assert.Equal(t, "00FF00", fill.SolidRGBColor.Val)

	// the fill is preserved on the round trip
	out := filepath.Join(t.TempDir(), "fill.pptx")
	assert.NoError(t, f.SaveAs(out))
	reopened, err := OpenFile(out)
	assert.NoError(t, err)
	defer reopened.Close()
	fill, err = reopened.GetShapeFill(257, 8)
	assert.NoError(t, err)
	assert.NotNil(t, fill)
	assert.Equal(t, "00FF00", fill.SolidRGBColor.Val)

	// the deleted fill is replaced by the noFill property
	err = reopened.DeleteShapeFill(257, 8)
	assert.NoError(t, err)
	fill, err = reopened.GetShapeFill(257, 8)
	assert.NoError(t, err)
	assert.Nil(t, fill)
}

// TestSetShapeXfrm checks the set of the position and the size (a:xfrm) of
// the shape by the slide id and the shape id: the xfrm is replaced, the xfrm
// is preserved on the round trip, the shape that is not in the slide returns
// the error.
func TestSetShapeXfrm(t *testing.T) {
	f, err := OpenFile(testFixture)
	assert.NoError(t, err)
	defer f.Close()

	// the shape that is not in the slide
	err = f.SetShapeXfrm(257, 999, DecodeXfrm{})
	assert.EqualError(t, err, ErrShapeNotExist{ShapeID: 999}.Error())
	// the slide that is not in the presentation
	err = f.SetShapeXfrm(999, 8, DecodeXfrm{})
	assert.EqualError(t, err, ErrSlideNotExist{SlideID: 999}.Error())

	// the xfrm is set to the shape
	xfrm := DecodeXfrm{
		Offset:  &Offset{X: 2000000, Y: 2000000},
		Extents: &Extents{CX: 4000000, CY: 2000000},
	}
	err = f.SetShapeXfrm(257, 8, xfrm)
	assert.NoError(t, err)

	shapes, err := f.GetShapes(257)
	assert.NoError(t, err)
	var shape *decodeShape
	for i := range shapes {
		if shapes[i].NonVisualShapeProperties.CommonNonVisualProperties.ID == 8 {
			shape = &shapes[i]
		}
	}
	assert.NotNil(t, shape)
	assert.Equal(t, 2000000, shape.ShapeProperties.Xfrm.Offset.X)
	assert.Equal(t, 4000000, shape.ShapeProperties.Xfrm.Extents.CX)

	// the xfrm is preserved on the round trip
	out := filepath.Join(t.TempDir(), "xfrm.pptx")
	assert.NoError(t, f.SaveAs(out))
	reopened, err := OpenFile(out)
	assert.NoError(t, err)
	defer reopened.Close()
	shapes, err = reopened.GetShapes(257)
	assert.NoError(t, err)
	for i := range shapes {
		if shapes[i].NonVisualShapeProperties.CommonNonVisualProperties.ID == 8 {
			assert.Equal(t, 2000000, shapes[i].ShapeProperties.Xfrm.Offset.X)
			assert.Equal(t, 4000000, shapes[i].ShapeProperties.Xfrm.Extents.CX)
		}
	}
}

// TestCreateShape checks the create of the shape by the slide id: the id of
// the created shape is unique in the whole shape tree of the slide, the text
// body of the created shape is kept, the created shape is preserved on the
// round trip, the slide that is not in the presentation returns the error.
func TestCreateShape(t *testing.T) {
	f, err := OpenFile(testFixture)
	assert.NoError(t, err)
	defer f.Close()

	// the slide that is not in the presentation
	_, err = f.CreateShape(999, DecodeShapeProperties{}, DecodeTextBody{})
	assert.EqualError(t, err, ErrSlideNotExist{SlideID: 999}.Error())

	shapeID, err := f.CreateShape(257, DecodeShapeProperties{
		Xfrm: &DecodeXfrm{
			Offset:  &Offset{X: 1000000, Y: 1000000},
			Extents: &Extents{CX: 3000000, CY: 2000000},
		},
	}, DecodeTextBody{
		BodyProperties: &DecodeBodyProperties{},
		ListStyle:      &struct{}{},
		Paragraph: []DecodeParagraph{
			{Runs: []DecodeRuns{{Text: "new shape"}}},
		},
	})
	assert.NoError(t, err)
	assert.Equal(t, 9, shapeID)

	shapes, err := f.GetShapes(257)
	assert.NoError(t, err)
	assert.Len(t, shapes, 3)
	var shape *decodeShape
	for i := range shapes {
		if shapes[i].NonVisualShapeProperties.CommonNonVisualProperties.ID == shapeID {
			shape = &shapes[i]
		}
	}
	assert.NotNil(t, shape)
	assert.Equal(t, "new shape", shape.TextBody.Paragraph[0].Runs[0].Text)

	// the created shape is preserved on the round trip
	out := filepath.Join(t.TempDir(), "shape.pptx")
	assert.NoError(t, f.SaveAs(out))
	reopened, err := OpenFile(out)
	assert.NoError(t, err)
	defer reopened.Close()
	shapes, err = reopened.GetShapes(257)
	assert.NoError(t, err)
	assert.Len(t, shapes, 3)
	for i := range shapes {
		if shapes[i].NonVisualShapeProperties.CommonNonVisualProperties.ID == shapeID {
			assert.Equal(t, "new shape", shapes[i].TextBody.Paragraph[0].Runs[0].Text)
			assert.Equal(t, 1000000, shapes[i].ShapeProperties.Xfrm.Offset.X)
		}
	}
}

// TestDeleteShape checks the delete of the shape by the slide id and the
// shape id: the shape that is not in the slide returns the error, the add
// and the delete of the shape keep the slide unchanged on the round trip.
func TestDeleteShape(t *testing.T) {
	f, err := OpenFile(testFixture)
	assert.NoError(t, err)
	defer f.Close()

	// the shape that is not in the slide
	err = f.DeleteShape(257, 999)
	assert.EqualError(t, err, ErrShapeNotExist{ShapeID: 999}.Error())
	// the slide that is not in the presentation
	err = f.DeleteShape(999, 8)
	assert.EqualError(t, err, ErrSlideNotExist{SlideID: 999}.Error())

	// the add and the delete of the shape keep the slide unchanged on the
	// round trip
	shapeID, err := f.CreateShape(257, DecodeShapeProperties{}, DecodeTextBody{})
	assert.NoError(t, err)
	err = f.DeleteShape(257, shapeID)
	assert.NoError(t, err)

	shapes, err := f.GetShapes(257)
	assert.NoError(t, err)
	assert.Len(t, shapes, 2)

	out := filepath.Join(t.TempDir(), "shape.pptx")
	assert.NoError(t, f.SaveAs(out))
	assert.NoError(t, assertPartsEqual(t, testFixture, out))
}

// TestSetShapeTextMultipleLines checks the set of the text of the shape by
// the slide id and the shape id: the text is set as a single paragraph with
// one run per line and a line break between the lines, the text is preserved
// on the round trip, the shape that is not in the slide returns the error.
func TestSetShapeTextMultipleLines(t *testing.T) {
	f, err := OpenFile(testFixture)
	assert.NoError(t, err)
	defer f.Close()

	// the shape that is not in the slide
	err = f.SetShapeText(257, 999, "text")
	assert.EqualError(t, err, ErrShapeNotExist{ShapeID: 999}.Error())
	// the slide that is not in the presentation
	err = f.SetShapeText(999, 8, "text")
	assert.EqualError(t, err, ErrSlideNotExist{SlideID: 999}.Error())

	// the text is set to the shape
	err = f.SetShapeText(257, 8, "Line 1\nLine 2")
	assert.NoError(t, err)

	shapes, err := f.GetShapes(257)
	assert.NoError(t, err)
	var shape *decodeShape
	for i := range shapes {
		if shapes[i].NonVisualShapeProperties.CommonNonVisualProperties.ID == 8 {
			shape = &shapes[i]
		}
	}
	assert.NotNil(t, shape)
	paragraph := shape.TextBody.Paragraph[0]
	assert.Len(t, shape.TextBody.Paragraph, 1)
	assert.Equal(t, "Line 1", paragraph.Runs[0].Text)
	assert.Equal(t, "Line 2", paragraph.Runs[1].Text)
	assert.Len(t, paragraph.Breaks, 1)

	// the text is preserved on the round trip
	out := filepath.Join(t.TempDir(), "text.pptx")
	assert.NoError(t, f.SaveAs(out))
	reopened, err := OpenFile(out)
	assert.NoError(t, err)
	defer reopened.Close()
	shapes, err = reopened.GetShapes(257)
	assert.NoError(t, err)
	for i := range shapes {
		if shapes[i].NonVisualShapeProperties.CommonNonVisualProperties.ID == 8 {
			assert.Equal(t, "Line 1", shapes[i].TextBody.Paragraph[0].Runs[0].Text)
			assert.Equal(t, "Line 2", shapes[i].TextBody.Paragraph[0].Runs[1].Text)
		}
	}
}

// TestSetShapeTextBody checks the set of the text body of the shape by the
// slide id and the shape id: the text body is replaced, the text body is
// preserved on the round trip, the shape that is not in the slide returns
// the error.
func TestSetShapeTextBody(t *testing.T) {
	f, err := OpenFile(testFixture)
	assert.NoError(t, err)
	defer f.Close()

	// the shape that is not in the slide
	err = f.SetShapeTextBody(257, 999, DecodeTextBody{})
	assert.EqualError(t, err, ErrShapeNotExist{ShapeID: 999}.Error())
	// the slide that is not in the presentation
	err = f.SetShapeTextBody(999, 8, DecodeTextBody{})
	assert.EqualError(t, err, ErrSlideNotExist{SlideID: 999}.Error())

	// the text body is replaced
	err = f.SetShapeTextBody(257, 8, DecodeTextBody{
		BodyProperties: &DecodeBodyProperties{},
		ListStyle:      &struct{}{},
		Paragraph: []DecodeParagraph{
			{Runs: []DecodeRuns{{Text: "replaced"}}},
		},
	})
	assert.NoError(t, err)

	shapes, err := f.GetShapes(257)
	assert.NoError(t, err)
	var shape *decodeShape
	for i := range shapes {
		if shapes[i].NonVisualShapeProperties.CommonNonVisualProperties.ID == 8 {
			shape = &shapes[i]
		}
	}
	assert.NotNil(t, shape)
	assert.Equal(t, "replaced", shape.TextBody.Paragraph[0].Runs[0].Text)

	// the text body is preserved on the round trip
	out := filepath.Join(t.TempDir(), "text.pptx")
	assert.NoError(t, f.SaveAs(out))
	reopened, err := OpenFile(out)
	assert.NoError(t, err)
	defer reopened.Close()
	shapes, err = reopened.GetShapes(257)
	assert.NoError(t, err)
	for i := range shapes {
		if shapes[i].NonVisualShapeProperties.CommonNonVisualProperties.ID == 8 {
			assert.Equal(t, "replaced", shapes[i].TextBody.Paragraph[0].Runs[0].Text)
		}
	}
}

// TestSetShapeOrder checks the set of the order (the z-order) of the
// children of the shape tree of the slide: the order is changed, the order
// is preserved on the round trip, the order that is not the permutation of
// the children of the shape tree returns the error.
func TestSetShapeOrder(t *testing.T) {
	f, err := OpenFile(testFixture)
	assert.NoError(t, err)
	defer f.Close()

	// the slide that is not in the presentation
	err = f.SetShapeOrder(999, nil)
	assert.EqualError(t, err, ErrSlideNotExist{SlideID: 999}.Error())

	// the order of the shape tree of the empty slide
	shapeOrder, err := f.GetShapeOrder(257)
	assert.NoError(t, err)
	assert.Equal(t, []ShapeOrder{
		{Kind: ShapeOrderSp, ID: 7},
		{Kind: ShapeOrderSp, ID: 8},
	}, shapeOrder)

	// the order that is not the permutation of the children of the shape
	// tree
	err = f.SetShapeOrder(257, []ShapeOrder{{Kind: ShapeOrderSp, ID: 7}})
	assert.EqualError(t, err, ErrShapeOrderInvalid.Error())
	err = f.SetShapeOrder(257, []ShapeOrder{{Kind: ShapeOrderSp, ID: 7}, {Kind: ShapeOrderSp, ID: 8}, {Kind: ShapeOrderSp, ID: 8}})
	assert.EqualError(t, err, ErrShapeOrderInvalid.Error())

	// the order is changed
	err = f.SetShapeOrder(257, []ShapeOrder{
		{Kind: ShapeOrderSp, ID: 8},
		{Kind: ShapeOrderSp, ID: 7},
	})
	assert.NoError(t, err)
	shapeOrder, err = f.GetShapeOrder(257)
	assert.NoError(t, err)
	assert.Equal(t, []ShapeOrder{
		{Kind: ShapeOrderSp, ID: 8},
		{Kind: ShapeOrderSp, ID: 7},
	}, shapeOrder)

	// the order is preserved on the round trip
	out := filepath.Join(t.TempDir(), "order.pptx")
	assert.NoError(t, f.SaveAs(out))
	slideXML := string(readZipPart(t, out, "ppt/slides/slide2.xml"))
	assert.Less(t, strings.Index(slideXML, `id="8"`), strings.Index(slideXML, `id="7"`))
}

// TestMoveShape checks the move of the shape to a specified position in the
// shape tree of the slide: the shape with the source id is moved before the
// shape with the target id, the shape that is not in the slide returns the
// error, the moved order is preserved on the round trip.
func TestMoveShape(t *testing.T) {
	f, err := OpenFile(testFixture)
	assert.NoError(t, err)
	defer f.Close()

	// the slide that is not in the presentation
	err = f.MoveShape(999, 7, 8)
	assert.EqualError(t, err, ErrSlideNotExist{SlideID: 999}.Error())
	// the shape that is not in the slide
	err = f.MoveShape(257, 999, 8)
	assert.EqualError(t, err, ErrShapeNotExist{999}.Error())
	err = f.MoveShape(257, 7, 999)
	assert.EqualError(t, err, ErrShapeNotExist{999}.Error())

	// the shape is moved before the target shape
	err = f.MoveShape(257, 8, 7)
	assert.NoError(t, err)
	shapeOrder, err := f.GetShapeOrder(257)
	assert.NoError(t, err)
	assert.Equal(t, []ShapeOrder{
		{Kind: ShapeOrderSp, ID: 8},
		{Kind: ShapeOrderSp, ID: 7},
	}, shapeOrder)

	// the moved shape is in the new position after the round trip
	out := filepath.Join(t.TempDir(), "order.pptx")
	assert.NoError(t, f.SaveAs(out))
	slideXML := string(readZipPart(t, out, "ppt/slides/slide2.xml"))
	assert.Less(t, strings.Index(slideXML, `id="8"`), strings.Index(slideXML, `id="7"`))
}

// TestGetGroupShapeProperties checks the get of the group shape properties
// (the xfrm of the shape tree of the slide) and the non-visual group shape
// properties (the cNvPr of the shape tree of the slide) by the slide id,
// the slide that is not in the presentation returns the error.
func TestGetGroupShapeProperties(t *testing.T) {
	f, err := OpenFile(testFixture)
	assert.NoError(t, err)
	defer f.Close()

	groupShapeProperties, err := f.GetGroupShapeProperties(257)
	assert.NoError(t, err)
	assert.NotNil(t, groupShapeProperties)
	assert.Equal(t, 0, groupShapeProperties.Xfrm.Offset.X)
	assert.Equal(t, 0, groupShapeProperties.Xfrm.Extents.CX)
	assert.Equal(t, 0, groupShapeProperties.Xfrm.ChildOffset.X)
	assert.Equal(t, 0, groupShapeProperties.Xfrm.ChildExtents.CY)

	nonVisual, err := f.GetNonVisualGroupShapeProperties(257)
	assert.NoError(t, err)
	assert.NotNil(t, nonVisual)
	assert.Equal(t, 1, nonVisual.CommonNonVisualProperties.ID)

	// the slide that is not in the presentation
	_, err = f.GetGroupShapeProperties(999)
	assert.EqualError(t, err, ErrSlideNotExist{SlideID: 999}.Error())
	_, err = f.GetNonVisualGroupShapeProperties(999)
	assert.EqualError(t, err, ErrSlideNotExist{SlideID: 999}.Error())
}

// TestShapeOrderKinds checks that the shape order contains the entries of
// every kind of the children of the shape tree of the slide: the shapes,
// the pictures, the connectors and the graphic frames (the tables, the
// charts), in the document order.
func TestShapeOrderKinds(t *testing.T) {
	// the pictures of the slide (the slide 258 of the test fixture has two
	// pictures)
	f, err := OpenFile(testFixture)
	assert.NoError(t, err)
	defer f.Close()
	shapeOrder, err := f.GetShapeOrder(258)
	assert.NoError(t, err)
	assert.Len(t, shapeOrder, 4)
	var hasPic bool
	for _, entry := range shapeOrder {
		if entry.Kind == ShapeOrderPic {
			hasPic = true
		}
	}
	assert.True(t, hasPic, "the pictures are not in the shape order")

	// the connector of the slide
	connectorID, err := f.CreateConnector(257, DecodeShapeProperties{})
	assert.NoError(t, err)
	shapeOrder, err = f.GetShapeOrder(257)
	assert.NoError(t, err)
	assert.Contains(t, shapeOrder, ShapeOrder{Kind: ShapeOrderCxnSp, ID: connectorID})

	// the graphic frame of the chart (the slide 256 of the chart fixture)
	cf, err := OpenFile(filepath.Join("test", "chart.pptx"))
	assert.NoError(t, err)
	defer cf.Close()
	charts, err := cf.GetCharts(256)
	assert.NoError(t, err)
	assert.Len(t, charts, 1)
	shapeOrder, err = cf.GetShapeOrder(256)
	assert.NoError(t, err)
	assert.Contains(t, shapeOrder, ShapeOrder{Kind: ShapeOrderGraphicFrame, ID: frameID(charts[0])})
}

// TestGetSlideList checks the get of the slides of the presentation and the
// get of the slide index of the presentation by the slide id, the slide that
// is not in the presentation returns the index -1.
func TestGetSlideList(t *testing.T) {
	f, err := OpenFile(testFixture)
	assert.NoError(t, err)
	defer f.Close()

	assert.Equal(t, []int{256, 257, 258}, f.GetSlideList())

	idx, err := f.GetSlideIndex(257)
	assert.NoError(t, err)
	assert.Equal(t, 1, idx)

	idx, err = f.GetSlideIndex(999)
	assert.NoError(t, err)
	assert.Equal(t, -1, idx)

	assert.Equal(t, 0, f.GetActiveSlideIndex())
}

// TestDeleteSlide checks the delete of the slide by the slide id: the slide
// is removed from the slide list of the presentation, the slide part and
// the relationships of the slide are removed, the remaining slides are
// preserved on the round trip.
func TestDeleteSlide(t *testing.T) {
	f, err := OpenFile(testFixture)
	assert.NoError(t, err)
	defer f.Close()

	assert.Equal(t, []int{256, 257, 258}, f.GetSlideList())

	// the slide that is not in the presentation
	assert.NoError(t, f.DeleteSlide(999))
	assert.Equal(t, []int{256, 257, 258}, f.GetSlideList())

	// the slide is deleted
	assert.NoError(t, f.DeleteSlide(258))
	assert.Equal(t, []int{256, 257}, f.GetSlideList())

	out := filepath.Join(t.TempDir(), "delete_slide.pptx")
	assert.NoError(t, f.SaveAs(out))

	// the slide part of the deleted slide is not in the saved file
	assert.False(t, hasZipPart(t, out, "ppt/slides/slide3.xml"))

	reopened, err := OpenFile(out)
	assert.NoError(t, err)
	defer reopened.Close()
	assert.Equal(t, []int{256, 257}, reopened.GetSlideList())
}

// TestSave checks the save of the presentation: the save without the path
// returns the error, the save to the path keeps the content of the
// presentation.
func TestSave(t *testing.T) {
	// the save without the path
	f := NewFile()
	assert.EqualError(t, f.Save(), ErrSave.Error())

	f, err := OpenFile(testFixture)
	assert.NoError(t, err)
	defer f.Close()

	out := filepath.Join(t.TempDir(), "save.pptx")
	assert.NoError(t, f.SaveAs(out))
	assert.NoError(t, assertPartsEqual(t, testFixture, out))
}
