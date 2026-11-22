// Copyright 2026 kenny-not-dead. All rights reserved.
// Use of this source code is governed by a BSD-style license
// that can be found in the LICENSE file.
//
// Package gopptx provides functionality to create and manipulate PowerPoint
// (.pptx) files in Go, using the Office Open XML (ECMA-376) format.

package gopptx

import (
	"bytes"
	"errors"
	"path/filepath"
	"strconv"
	"strings"
	"testing"
)

// slideMasterTestID is the id of the slide master of the new file (the id
// attribute of the p:sldMasterId element of the template presentation).
const slideMasterTestID = 2147483648

// slideLayoutTestID is the id of the slide layout of the new file (the id
// attribute of the p:sldLayoutId element of the template slide master).
const slideLayoutTestID = 2147483649

// slideMasterTestRelsPath is the part path of the slide master rels of the
// new file.
const slideMasterTestRelsPath = "ppt/slideMasters/_rels/slideMaster1.xml.rels"

// TestGetSlideLayoutListAndMasterList checks that the slide layout list and
// the slide master list of the new file are read from the slide master and
// the presentation parts.
func TestGetSlideLayoutListAndMasterList(t *testing.T) {
	f := NewFile()
	defer f.Close()

	if list := f.GetSlideLayoutList(); len(list) != 1 || list[0] != slideLayoutTestID {
		t.Fatalf("GetSlideLayoutList = %v, want [%d]", list, slideLayoutTestID)
	}
	if list := f.GetSlideMasterList(); len(list) != 1 || list[0] != slideMasterTestID {
		t.Fatalf("GetSlideMasterList = %v, want [%d]", list, slideMasterTestID)
	}

	// the shapes of the layout and the master are read from the parts
	shapes, err := f.GetLayoutShapes(slideLayoutTestID)
	if err != nil {
		t.Fatal(err)
	}
	if len(shapes) == 0 {
		t.Fatal("GetLayoutShapes returned no shapes")
	}
	shapes, err = f.GetMasterShapes(slideMasterTestID)
	if err != nil {
		t.Fatal(err)
	}
	if len(shapes) == 0 {
		t.Fatal("GetMasterShapes returned no shapes")
	}

	// the non existing layout and master
	if _, err := f.GetLayoutShapes(1); !errors.As(err, new(ErrSlideLayoutNotExist)) {
		t.Fatalf("expected ErrSlideLayoutNotExist, got %v", err)
	}
	if _, err := f.GetMasterShapes(1); !errors.As(err, new(ErrSlideMasterNotExist)) {
		t.Fatalf("expected ErrSlideMasterNotExist, got %v", err)
	}
}

// TestLayoutShapeCreateDelete checks that CreateLayoutShape creates the
// shape in the shape tree of the slide layout (the id of the shape is the
// next free id of the layout) and DeleteLayoutShape deletes it.
func TestLayoutShapeCreateDelete(t *testing.T) {
	f := NewFile()
	defer f.Close()

	shapeID, err := f.CreateLayoutShape(slideLayoutTestID, DecodeShapeProperties{
		Xfrm: &DecodeXfrm{
			Offset:  &Offset{X: 457200, Y: 457200},
			Extents: &Extents{CX: 914400, CY: 457200},
		},
		PresetGeometry: &DecodePresetGeometry{Preset: "rect"},
	}, DecodeTextBody{Paragraph: []DecodeParagraph{{}}})
	if err != nil {
		t.Fatal(err)
	}
	// the max id of the shapes of the template layout is 6
	if shapeID != 7 {
		t.Fatalf("CreateLayoutShape id = %d, want 7", shapeID)
	}

	out := filepath.Join(t.TempDir(), "layout_shape.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	layout := string(readZipPart(t, out, layoutTestPath))
	if !strings.Contains(layout, `<p:cNvPr id="7" name=""/>`) || !strings.Contains(layout, `prst="rect"`) {
		t.Fatalf("the new shape is not in the layout:\n%s", layout)
	}

	if err := f.DeleteLayoutShape(slideLayoutTestID, shapeID); err != nil {
		t.Fatal(err)
	}
	if err := f.DeleteLayoutShape(slideLayoutTestID, shapeID); !errors.As(err, new(ErrShapeNotExist)) {
		t.Fatalf("expected ErrShapeNotExist, got %v", err)
	}
	shapes, err := f.GetLayoutShapes(slideLayoutTestID)
	if err != nil {
		t.Fatal(err)
	}
	for _, s := range shapes {
		if s.NonVisualShapeProperties != nil &&
			s.NonVisualShapeProperties.CommonNonVisualProperties != nil &&
			s.NonVisualShapeProperties.CommonNonVisualProperties.ID == shapeID {
			t.Fatalf("the shape %d is not deleted from the layout", shapeID)
		}
	}
}

// TestLayoutShapeText checks that SetLayoutShapeText and
// SetLayoutShapeTextBody set the text of the shape of the slide layout.
func TestLayoutShapeText(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// the title placeholder of the template layout
	shapes, err := f.GetLayoutShapes(slideLayoutTestID)
	if err != nil {
		t.Fatal(err)
	}
	var titleID int
	for _, s := range shapes {
		if s.NonVisualShapeProperties != nil &&
			s.NonVisualShapeProperties.NonVisualProperties != nil &&
			s.NonVisualShapeProperties.NonVisualProperties.Ph != nil &&
			s.NonVisualShapeProperties.NonVisualProperties.Ph.Type != nil &&
			*s.NonVisualShapeProperties.NonVisualProperties.Ph.Type == "title" {
			titleID = s.NonVisualShapeProperties.CommonNonVisualProperties.ID
		}
	}
	if titleID == 0 {
		t.Fatal("the title placeholder of the template layout is not found")
	}

	if err := f.SetLayoutShapeText(slideLayoutTestID, titleID, "Line 1\nLine 2"); err != nil {
		t.Fatal(err)
	}

	out := filepath.Join(t.TempDir(), "layout_text.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	layout := string(readZipPart(t, out, layoutTestPath))
	if !strings.Contains(layout, "<a:t>Line 1</a:t>") || !strings.Contains(layout, "<a:br/>") ||
		!strings.Contains(layout, "<a:t>Line 2</a:t>") {
		t.Fatalf("the text of the title placeholder is not set:\n%s", layout)
	}

	// the text body is replaced
	if err := f.SetLayoutShapeTextBody(slideLayoutTestID, titleID, DecodeTextBody{
		BodyProperties: &DecodeBodyProperties{},
		Paragraph:      []DecodeParagraph{{Runs: []DecodeRuns{{Text: "The new text"}}}},
	}); err != nil {
		t.Fatal(err)
	}
	shapes, _ = f.GetLayoutShapes(slideLayoutTestID)
	for _, s := range shapes {
		if s.NonVisualShapeProperties == nil ||
			s.NonVisualShapeProperties.CommonNonVisualProperties == nil ||
			s.NonVisualShapeProperties.CommonNonVisualProperties.ID != titleID ||
			s.NonVisualShapeProperties.NonVisualProperties == nil ||
			s.NonVisualShapeProperties.NonVisualProperties.Ph == nil ||
			s.NonVisualShapeProperties.NonVisualProperties.Ph.Type == nil ||
			*s.NonVisualShapeProperties.NonVisualProperties.Ph.Type != "title" {
			// the template layout has the duplicated shape ids (the title
			// and the slide number placeholder share the id): only the
			// title placeholder is checked
			continue
		}
		if len(s.TextBody.Paragraph) != 1 || len(s.TextBody.Paragraph[0].Runs) != 1 ||
			s.TextBody.Paragraph[0].Runs[0].Text != "The new text" {
			t.Fatalf("the text body of the title placeholder is not replaced")
		}
	}
}

// TestLayoutBackground checks that SetSlideLayoutBackground /
// GetSlideLayoutBackground / DeleteSlideLayoutBackground work with the
// background of the slide layout.
func TestLayoutBackground(t *testing.T) {
	f := NewFile()
	defer f.Close()

	bg := DecodeBackground{BGProperties: &DecodeBGProperties{
		SolidFill: &DecodeSolidFill{SolidRGBColor: &SolidRGBColor{Val: "F0F1F3"}},
	}}
	if err := f.SetSlideLayoutBackground(slideLayoutTestID, bg); err != nil {
		t.Fatal(err)
	}
	got, err := f.GetSlideLayoutBackground(slideLayoutTestID)
	if err != nil {
		t.Fatal(err)
	}
	if got == nil || got.BGProperties == nil || got.BGProperties.SolidFill == nil ||
		got.BGProperties.SolidFill.SolidRGBColor == nil ||
		got.BGProperties.SolidFill.SolidRGBColor.Val != "F0F1F3" {
		t.Fatalf("GetSlideLayoutBackground = %+v", got)
	}

	out := filepath.Join(t.TempDir(), "layout_bg.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	layout := string(readZipPart(t, out, layoutTestPath))
	if !strings.Contains(layout, `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="F0F1F3"/></a:solidFill></p:bgPr></p:bg>`) {
		t.Fatalf("the background of the layout is not set:\n%s", layout)
	}

	if err := f.DeleteSlideLayoutBackground(slideLayoutTestID); err != nil {
		t.Fatal(err)
	}
	got, _ = f.GetSlideLayoutBackground(slideLayoutTestID)
	if got != nil {
		t.Fatalf("the background of the layout is not deleted: %+v", got)
	}
}

// TestLayoutShapeXfrmFill checks that SetLayoutShapeXfrm /
// SetLayoutShapeFill / GetLayoutShapeFill work with the shapes of the slide
// layout.
func TestLayoutShapeXfrmFill(t *testing.T) {
	f := NewFile()
	defer f.Close()

	shapeID, err := f.CreateLayoutShape(slideLayoutTestID, DecodeShapeProperties{
		Xfrm: &DecodeXfrm{
			Offset:  &Offset{X: 0, Y: 0},
			Extents: &Extents{CX: 914400, CY: 457200},
		},
	}, DecodeTextBody{Paragraph: []DecodeParagraph{{}}})
	if err != nil {
		t.Fatal(err)
	}

	xfrm := DecodeXfrm{
		Offset:  &Offset{X: 8709025, Y: 4756150},
		Extents: &Extents{CX: 914400, CY: 457200},
	}
	if err := f.SetLayoutShapeXfrm(slideLayoutTestID, shapeID, xfrm); err != nil {
		t.Fatal(err)
	}
	if err := f.SetLayoutShapeFill(slideLayoutTestID, shapeID, DecodeSolidFill{
		SolidRGBColor: &SolidRGBColor{Val: "178A8A"},
	}); err != nil {
		t.Fatal(err)
	}
	got, err := f.GetLayoutShapeFill(slideLayoutTestID, shapeID)
	if err != nil {
		t.Fatal(err)
	}
	if got == nil || got.SolidRGBColor == nil || got.SolidRGBColor.Val != "178A8A" {
		t.Fatalf("GetLayoutShapeFill = %+v", got)
	}

	out := filepath.Join(t.TempDir(), "layout_xfrm_fill.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	layout := string(readZipPart(t, out, layoutTestPath))
	if !strings.Contains(layout, `<a:off x="8709025" y="4756150"/>`) ||
		!strings.Contains(layout, `<a:srgbClr val="178A8A"/>`) {
		t.Fatalf("the xfrm or the fill of the shape is not set:\n%s", layout)
	}
}

// TestMasterShapeCreateDelete checks that CreateMasterShape creates the
// shape in the shape tree of the slide master (the id of the shape is the
// next free id of the master) and DeleteMasterShape deletes it.
func TestMasterShapeCreateDelete(t *testing.T) {
	f := NewFile()
	defer f.Close()

	shapeID, err := f.CreateMasterShape(slideMasterTestID, DecodeShapeProperties{
		Xfrm: &DecodeXfrm{
			Offset:  &Offset{X: 457200, Y: 457200},
			Extents: &Extents{CX: 914400, CY: 457200},
		},
		PresetGeometry: &DecodePresetGeometry{Preset: "rect"},
	}, DecodeTextBody{Paragraph: []DecodeParagraph{{Runs: []DecodeRuns{{Text: "The master logo text"}}}}})
	if err != nil {
		t.Fatal(err)
	}
	// the max id of the shapes of the template master is 4
	if shapeID != 5 {
		t.Fatalf("CreateMasterShape id = %d, want 5", shapeID)
	}

	out := filepath.Join(t.TempDir(), "master_shape.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	master := string(readZipPart(t, out, slideMasterTestPath))
	if !strings.Contains(master, `<p:cNvPr id="5" name=""/>`) ||
		!strings.Contains(master, "The master logo text") {
		t.Fatalf("the new shape is not in the master:\n%s", master)
	}
	// the elements of the master other than the cSld element are kept (the
	// template master has no the txStyles element)
	if !strings.Contains(master, `<p:clrMap bg1="lt1" bg2="lt2" tx1="dk1" tx2="dk2"`) ||
		!strings.Contains(master, `<p:sldLayoutId id="2147483649" r:id="rId2"/>`) {
		t.Fatalf("the elements of the master are lost:\n%s", master)
	}

	if err := f.DeleteMasterShape(slideMasterTestID, shapeID); err != nil {
		t.Fatal(err)
	}
	if err := f.DeleteMasterShape(slideMasterTestID, shapeID); !errors.As(err, new(ErrShapeNotExist)) {
		t.Fatalf("expected ErrShapeNotExist, got %v", err)
	}
}

// TestMasterShapeText checks that SetMasterShapeText and
// SetMasterShapeTextBody set the text of the shape of the slide master.
func TestMasterShapeText(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// the title placeholder of the template master (the id of the title
	// placeholder of the template master is 0)
	shapes, err := f.GetMasterShapes(slideMasterTestID)
	if err != nil {
		t.Fatal(err)
	}
	titleID := -1
	for _, s := range shapes {
		if s.NonVisualShapeProperties != nil &&
			s.NonVisualShapeProperties.NonVisualProperties != nil &&
			s.NonVisualShapeProperties.NonVisualProperties.Ph != nil &&
			s.NonVisualShapeProperties.NonVisualProperties.Ph.Type != nil &&
			*s.NonVisualShapeProperties.NonVisualProperties.Ph.Type == "title" {
			titleID = s.NonVisualShapeProperties.CommonNonVisualProperties.ID
		}
	}
	if titleID == -1 {
		t.Fatal("the title placeholder of the template master is not found")
	}

	if err := f.SetMasterShapeText(slideMasterTestID, titleID, "The master title"); err != nil {
		t.Fatal(err)
	}
	out := filepath.Join(t.TempDir(), "master_text.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	master := string(readZipPart(t, out, slideMasterTestPath))
	if !strings.Contains(master, "<a:t>The master title</a:t>") {
		t.Fatalf("the text of the title placeholder of the master is not set:\n%s", master)
	}

	if err := f.SetMasterShapeTextBody(slideMasterTestID, titleID, DecodeTextBody{
		BodyProperties: &DecodeBodyProperties{},
		Paragraph:      []DecodeParagraph{{Runs: []DecodeRuns{{Text: "The replaced text"}}}},
	}); err != nil {
		t.Fatal(err)
	}
	shapes, _ = f.GetMasterShapes(slideMasterTestID)
	for _, s := range shapes {
		if s.NonVisualShapeProperties == nil ||
			s.NonVisualShapeProperties.CommonNonVisualProperties == nil ||
			s.NonVisualShapeProperties.CommonNonVisualProperties.ID != titleID ||
			s.NonVisualShapeProperties.NonVisualProperties == nil ||
			s.NonVisualShapeProperties.NonVisualProperties.Ph == nil ||
			s.NonVisualShapeProperties.NonVisualProperties.Ph.Type == nil ||
			*s.NonVisualShapeProperties.NonVisualProperties.Ph.Type != "title" {
			continue
		}
		if len(s.TextBody.Paragraph) != 1 || s.TextBody.Paragraph[0].Runs[0].Text != "The replaced text" {
			t.Fatalf("the text body of the title placeholder of the master is not replaced")
		}
	}
}

// TestMasterBackground checks that SetSlideMasterBackground /
// GetSlideMasterBackground / DeleteSlideMasterBackground work with the
// background of the slide master.
func TestMasterBackground(t *testing.T) {
	f := NewFile()
	defer f.Close()

	bg := DecodeBackground{BGProperties: &DecodeBGProperties{
		SolidFill: &DecodeSolidFill{SolidRGBColor: &SolidRGBColor{Val: "0563C1"}},
	}}
	if err := f.SetSlideMasterBackground(slideMasterTestID, bg); err != nil {
		t.Fatal(err)
	}
	got, err := f.GetSlideMasterBackground(slideMasterTestID)
	if err != nil {
		t.Fatal(err)
	}
	if got == nil || got.BGProperties == nil || got.BGProperties.SolidFill == nil ||
		got.BGProperties.SolidFill.SolidRGBColor == nil ||
		got.BGProperties.SolidFill.SolidRGBColor.Val != "0563C1" {
		t.Fatalf("GetSlideMasterBackground = %+v", got)
	}

	out := filepath.Join(t.TempDir(), "master_bg.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	master := string(readZipPart(t, out, slideMasterTestPath))
	if !strings.Contains(master, `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="0563C1"/></a:solidFill></p:bgPr></p:bg>`) {
		t.Fatalf("the background of the master is not set:\n%s", master)
	}

	if err := f.DeleteSlideMasterBackground(slideMasterTestID); err != nil {
		t.Fatal(err)
	}
	got, _ = f.GetSlideMasterBackground(slideMasterTestID)
	if got != nil {
		t.Fatalf("the background of the master is not deleted: %+v", got)
	}
}

// TestLayoutShapeXfrmMaster checks that SetMasterShapeXfrm /
// SetMasterShapeFill / GetMasterShapeFill work with the shapes of the slide
// master.
func TestMasterShapeXfrmFill(t *testing.T) {
	f := NewFile()
	defer f.Close()

	shapeID, err := f.CreateMasterShape(slideMasterTestID, DecodeShapeProperties{
		Xfrm: &DecodeXfrm{
			Offset:  &Offset{X: 0, Y: 0},
			Extents: &Extents{CX: 914400, CY: 457200},
		},
	}, DecodeTextBody{Paragraph: []DecodeParagraph{{}}})
	if err != nil {
		t.Fatal(err)
	}

	xfrm := DecodeXfrm{
		Offset:  &Offset{X: 9165025, Y: 5213350},
		Extents: &Extents{CX: 914400, CY: 457200},
	}
	if err := f.SetMasterShapeXfrm(slideMasterTestID, shapeID, xfrm); err != nil {
		t.Fatal(err)
	}
	if err := f.SetMasterShapeFill(slideMasterTestID, shapeID, DecodeSolidFill{
		SolidRGBColor: &SolidRGBColor{Val: "D24525"},
	}); err != nil {
		t.Fatal(err)
	}
	got, err := f.GetMasterShapeFill(slideMasterTestID, shapeID)
	if err != nil {
		t.Fatal(err)
	}
	if got == nil || got.SolidRGBColor == nil || got.SolidRGBColor.Val != "D24525" {
		t.Fatalf("GetMasterShapeFill = %+v", got)
	}

	out := filepath.Join(t.TempDir(), "master_xfrm_fill.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	master := string(readZipPart(t, out, slideMasterTestPath))
	if !strings.Contains(master, `<a:off x="9165025" y="5213350"/>`) ||
		!strings.Contains(master, `<a:srgbClr val="D24525"/>`) {
		t.Fatalf("the xfrm or the fill of the shape of the master is not set:\n%s", master)
	}
}

// TestNewSlideLayout checks that NewSlideLayout creates the new slide layout
// (the copy of the source layout) in the slide master: the part and the rels
// part of the new layout, the content types, the relationship of the new
// layout in the rels of the slide master and the entry of the new layout in
// the slide layout id list of the slide master.
func TestNewSlideLayout(t *testing.T) {
	f := NewFile()
	defer f.Close()

	newLayoutID, err := f.NewSlideLayout(slideLayoutTestID)
	if err != nil {
		t.Fatal(err)
	}
	if newLayoutID != slideLayoutTestID+1 {
		t.Fatalf("NewSlideLayout id = %d, want %d", newLayoutID, slideLayoutTestID+1)
	}

	out := filepath.Join(t.TempDir(), "new_layout.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	newLayoutPart := "ppt/slideLayouts/slideLayout2.xml"
	// the part of the new layout is the copy of the part of the source layout
	sourceLayout := readZipPart(t, out, layoutTestPath)
	if !bytes.Equal(readZipPart(t, out, newLayoutPart), sourceLayout) {
		t.Fatalf("the part of the new layout is not the copy of the source layout")
	}
	// the content types of the new layout and its rels part (the attribute
	// order of the Override element of the content types part)
	contentTypes := string(readZipPart(t, out, "[Content_Types].xml"))
	if !strings.Contains(contentTypes, `ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml" PartName="/ppt/slideLayouts/slideLayout2.xml"`) ||
		!strings.Contains(contentTypes, `ContentType="application/vnd.openxmlformats-package.relationships+xml" PartName="/ppt/slideLayouts/_rels/slideLayout2.xml.rels"`) {
		t.Fatalf("the content types of the new layout are not added:\n%s", contentTypes)
	}
	// the rels part of the new layout points to the slide master
	newLayoutRels := string(readZipPart(t, out, "ppt/slideLayouts/_rels/slideLayout2.xml.rels"))
	if !strings.Contains(newLayoutRels, `relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"`) {
		t.Fatalf("the rels part of the new layout is not created:\n%s", newLayoutRels)
	}
	// the relationship of the new layout in the rels of the slide master
	masterRels := string(readZipPart(t, out, slideMasterTestRelsPath))
	if !strings.Contains(masterRels, `relationships/slideLayout" Target="../slideLayouts/slideLayout2.xml"`) {
		t.Fatalf("the relationship of the new layout is not added to the rels of the slide master:\n%s", masterRels)
	}
	// the entry of the new layout in the slide layout id list of the master
	master := string(readZipPart(t, out, slideMasterTestPath))
	if !strings.Contains(master, `<p:sldLayoutId id="`+strconv.Itoa(newLayoutID)+`" r:id="rId3"/>`) {
		t.Fatalf("the entry of the new layout is not added to the slide layout id list of the master:\n%s", master)
	}
	// the list of the layouts of the presentation
	if list := f.GetSlideLayoutList(); len(list) != 2 || list[1] != newLayoutID {
		t.Fatalf("GetSlideLayoutList = %v, want [%d %d]", list, slideLayoutTestID, newLayoutID)
	}
	// the second new layout (the id is incremented again)
	secondLayoutID, err := f.NewSlideLayout(slideLayoutTestID)
	if err != nil {
		t.Fatal(err)
	}
	if secondLayoutID != slideLayoutTestID+2 {
		t.Fatalf("the second NewSlideLayout id = %d, want %d", secondLayoutID, slideLayoutTestID+2)
	}
	// the non existing source layout
	if _, err := f.NewSlideLayout(1); !errors.As(err, new(ErrSlideLayoutNotExist)) {
		t.Fatalf("expected ErrSlideLayoutNotExist, got %v", err)
	}
}

// TestLayoutMasterRoundTrip checks that the slide layout and the slide
// master modified by the layout and the master methods are stable on the
// next save: the file saved after the round trip is identical to the file
// saved right after the modification.
func TestLayoutMasterRoundTrip(t *testing.T) {
	f := NewFile()
	defer f.Close()

	bg := DecodeBackground{BGProperties: &DecodeBGProperties{
		SolidFill: &DecodeSolidFill{SolidRGBColor: &SolidRGBColor{Val: "F0F1F3"}},
	}}
	if err := f.SetSlideLayoutBackground(slideLayoutTestID, bg); err != nil {
		t.Fatal(err)
	}
	if err := f.SetSlideMasterBackground(slideMasterTestID, bg); err != nil {
		t.Fatal(err)
	}
	if _, err := f.CreateLayoutShape(slideLayoutTestID, DecodeShapeProperties{
		Xfrm: &DecodeXfrm{
			Offset:  &Offset{X: 457200, Y: 457200},
			Extents: &Extents{CX: 914400, CY: 457200},
		},
	}, DecodeTextBody{Paragraph: []DecodeParagraph{{}}}); err != nil {
		t.Fatal(err)
	}
	if _, err := f.CreateMasterShape(slideMasterTestID, DecodeShapeProperties{
		Xfrm: &DecodeXfrm{
			Offset:  &Offset{X: 457200, Y: 457200},
			Extents: &Extents{CX: 914400, CY: 457200},
		},
	}, DecodeTextBody{Paragraph: []DecodeParagraph{{}}}); err != nil {
		t.Fatal(err)
	}
	if _, err := f.NewSlideLayout(slideLayoutTestID); err != nil {
		t.Fatal(err)
	}

	first := filepath.Join(t.TempDir(), "layout_master_1.pptx")
	if err := f.SaveAs(first); err != nil {
		t.Fatal(err)
	}

	f2, err := OpenFile(first)
	if err != nil {
		t.Fatal(err)
	}
	defer f2.Close()

	second := filepath.Join(t.TempDir(), "layout_master_2.pptx")
	if err := f2.SaveAs(second); err != nil {
		t.Fatal(err)
	}
	for _, part := range []string{
		slideMasterTestPath, layoutTestPath, "ppt/slideLayouts/slideLayout2.xml",
		slideMasterTestRelsPath, "[Content_Types].xml",
	} {
		a := readZipPart(t, first, part)
		b := readZipPart(t, second, part)
		if !bytes.Equal(a, b) {
			t.Errorf("%s differs after the round trip:\n--- first\n%s\n--- second\n%s", part, a, b)
		}
	}
	// the ids of the layouts and the masters are read back
	if list := f2.GetSlideLayoutList(); len(list) != 2 {
		t.Fatalf("GetSlideLayoutList after the round trip = %v, want 2 layouts", list)
	}
	if list := f2.GetSlideMasterList(); len(list) != 1 || list[0] != slideMasterTestID {
		t.Fatalf("GetSlideMasterList after the round trip = %v", list)
	}
}

// TestLayoutUnsupportedShapeTree checks that the modification of the slide
// layout with the picture in the shape tree returns the
// ErrSlideLayoutUnsupported error.
func TestLayoutUnsupportedShapeTree(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// the slide layout with the picture in the shape tree
	orig := string(f.readBytes(layoutTestPath))
	withPic := strings.Replace(orig, "</p:spTree>",
		`<p:pic><p:nvPicPr><p:cNvPr id="100" name="pic"/></p:nvPicPr></p:pic></p:spTree>`, 1)
	f.Pkg.Store(layoutTestPath, []byte(withPic))

	_, err := f.GetLayoutShapes(slideLayoutTestID)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := f.CreateLayoutShape(slideLayoutTestID, DecodeShapeProperties{}, DecodeTextBody{}); !errors.As(err, new(ErrSlideLayoutUnsupported)) {
		t.Fatalf("expected ErrSlideLayoutUnsupported, got %v", err)
	}
	if err := f.SetSlideLayoutBackground(slideLayoutTestID, DecodeBackground{}); !errors.As(err, new(ErrSlideLayoutUnsupported)) {
		t.Fatalf("expected ErrSlideLayoutUnsupported, got %v", err)
	}
}

// TestMasterUnsupportedShapeTree checks that the modification of the slide
// master with the picture in the shape tree returns the
// ErrSlideMasterUnsupported error.
func TestMasterUnsupportedShapeTree(t *testing.T) {
	f := NewFile()
	defer f.Close()

	orig := string(f.readBytes(slideMasterTestPath))
	withPic := strings.Replace(orig, "</p:spTree>",
		`<p:pic><p:nvPicPr><p:cNvPr id="100" name="pic"/></p:nvPicPr></p:pic></p:spTree>`, 1)
	f.Pkg.Store(slideMasterTestPath, []byte(withPic))

	_, err := f.GetMasterShapes(slideMasterTestID)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := f.CreateMasterShape(slideMasterTestID, DecodeShapeProperties{}, DecodeTextBody{}); !errors.As(err, new(ErrSlideMasterUnsupported)) {
		t.Fatalf("expected ErrSlideMasterUnsupported, got %v", err)
	}
	if err := f.SetSlideMasterBackground(slideMasterTestID, DecodeBackground{}); !errors.As(err, new(ErrSlideMasterUnsupported)) {
		t.Fatalf("expected ErrSlideMasterUnsupported, got %v", err)
	}
}

// TestMasterColorMap checks that GetSlideMasterColorMap reads the color map
// (p:clrMap) of the slide master and SetSlideMasterColorMap sets the color
// map of the slide master (the color map is kept on the round trip).
func TestMasterColorMap(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// the template master has the identity color map (the attributes of the
	// color map are the names of the theme scheme colors)
	cm, err := f.GetSlideMasterColorMap(slideMasterTestID)
	if err != nil {
		t.Fatal(err)
	}
	if cm == nil || cm.BG1 != "lt1" || cm.BG2 != "lt2" || cm.TX1 != "dk1" || cm.TX2 != "dk2" ||
		cm.HLink != "hlink" || cm.FolHLink != "folHlink" || cm.Accent1 != "accent1" {
		t.Fatalf("GetSlideMasterColorMap = %+v", cm)
	}

	// the hyperlink color of the master is mapped to the accent1 scheme
	// color
	cm.HLink = "accent1"
	if err := f.SetSlideMasterColorMap(slideMasterTestID, *cm); err != nil {
		t.Fatal(err)
	}
	got, err := f.GetSlideMasterColorMap(slideMasterTestID)
	if err != nil {
		t.Fatal(err)
	}
	if got == nil || got.HLink != "accent1" {
		t.Fatalf("GetSlideMasterColorMap after the set = %+v", got)
	}

	out := filepath.Join(t.TempDir(), "master_clrmap.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	master := string(readZipPart(t, out, slideMasterTestPath))
	if !strings.Contains(master, `<p:clrMap bg1="lt1" bg2="lt2" tx1="dk1" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="accent1" folHlink="folHlink"/>`) {
		t.Fatalf("the color map of the master is not set:\n%s", master)
	}

	// the color map is kept on the round trip of the saved file
	f2, err := OpenFile(out)
	if err != nil {
		t.Fatal(err)
	}
	defer f2.Close()
	got, err = f2.GetSlideMasterColorMap(slideMasterTestID)
	if err != nil {
		t.Fatal(err)
	}
	if got == nil || got.HLink != "accent1" || got.FolHLink != "folHlink" {
		t.Fatalf("GetSlideMasterColorMap after the round trip = %+v", got)
	}

	// the non existing master
	if _, err := f.GetSlideMasterColorMap(1); !errors.As(err, new(ErrSlideMasterNotExist)) {
		t.Fatalf("expected ErrSlideMasterNotExist, got %v", err)
	}
}

// TestMasterTextStyles checks that GetSlideMasterTextStyles reads the text
// styles (p:txStyles) of the slide master, SetSlideMasterTextStyles sets the
// text styles of the slide master (the text styles are inherited by the
// slides) and DeleteSlideMasterTextStyles deletes them.
func TestMasterTextStyles(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// the template master has no text styles element
	ts, err := f.GetSlideMasterTextStyles(slideMasterTestID)
	if err != nil {
		t.Fatal(err)
	}
	if ts != nil {
		t.Fatalf("GetSlideMasterTextStyles = %+v, want nil", ts)
	}

	// the title style: the default run properties of the level 1 (the size
	// and the font), the body style: the bullet character of the level 1
	size := 3200
	textStyles := DecodeTextStyles{
		TitleStyle: &DecodeTextStyle{
			Lvl1: &DecodeParagraphProperties{
				DefRunProperties: &DecodeRunProperties{
					Size:  &size,
					Latin: &Font{Typeface: "Arial"},
				},
			},
		},
		BodyStyle: &DecodeTextStyle{
			Lvl1: &DecodeParagraphProperties{
				BuChar: &DecodeBulletChar{Char: "•"},
			},
		},
	}
	if err := f.SetSlideMasterTextStyles(slideMasterTestID, textStyles); err != nil {
		t.Fatal(err)
	}
	got, err := f.GetSlideMasterTextStyles(slideMasterTestID)
	if err != nil {
		t.Fatal(err)
	}
	if got == nil || got.TitleStyle == nil || got.TitleStyle.Lvl1 == nil ||
		got.TitleStyle.Lvl1.DefRunProperties == nil ||
		got.TitleStyle.Lvl1.DefRunProperties.Size == nil || *got.TitleStyle.Lvl1.DefRunProperties.Size != size ||
		got.BodyStyle == nil || got.BodyStyle.Lvl1 == nil || got.BodyStyle.Lvl1.BuChar == nil ||
		got.BodyStyle.Lvl1.BuChar.Char != "•" {
		t.Fatalf("GetSlideMasterTextStyles after the set = %+v", got)
	}

	out := filepath.Join(t.TempDir(), "master_txstyles.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	master := string(readZipPart(t, out, slideMasterTestPath))
	if !strings.Contains(master, "<p:txStyles><p:titleStyle><a:lvl1pPr>") ||
		!strings.Contains(master, `<a:defRPr sz="3200">`) ||
		!strings.Contains(master, `<a:latin typeface="Arial"/>`) ||
		!strings.Contains(master, "<p:bodyStyle><a:lvl1pPr><a:buChar char=\"•\"/></a:lvl1pPr></p:bodyStyle>") {
		t.Fatalf("the text styles of the master are not set:\n%s", master)
	}

	// the text styles are kept on the round trip of the saved file
	f2, err := OpenFile(out)
	if err != nil {
		t.Fatal(err)
	}
	defer f2.Close()
	got, err = f2.GetSlideMasterTextStyles(slideMasterTestID)
	if err != nil {
		t.Fatal(err)
	}
	if got == nil || got.TitleStyle == nil || got.TitleStyle.Lvl1 == nil ||
		got.TitleStyle.Lvl1.DefRunProperties == nil ||
		got.TitleStyle.Lvl1.DefRunProperties.Size == nil || *got.TitleStyle.Lvl1.DefRunProperties.Size != size {
		t.Fatalf("GetSlideMasterTextStyles after the round trip = %+v", got)
	}

	// the text styles are deleted
	if err := f.DeleteSlideMasterTextStyles(slideMasterTestID); err != nil {
		t.Fatal(err)
	}
	got, err = f.GetSlideMasterTextStyles(slideMasterTestID)
	if err != nil {
		t.Fatal(err)
	}
	if got != nil {
		t.Fatalf("the text styles of the master are not deleted: %+v", got)
	}
	out = filepath.Join(t.TempDir(), "master_txstyles_none.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	master = string(readZipPart(t, out, slideMasterTestPath))
	if strings.Contains(master, "<p:txStyles>") {
		t.Fatalf("the text styles element is not deleted from the master:\n%s", master)
	}
}

// TestMasterHeaderFooterPreserved checks that the attributes of the header
// footer (p:hf) element of the slide master are preserved on the master
// round trip (the element is modeled by the attributes, the same way the
// clrMap element is modeled).
func TestMasterHeaderFooterPreserved(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// the template master has no header footer element: it is injected
	orig := string(f.readBytes(slideMasterTestPath))
	withHF := strings.Replace(orig, "</p:sldMaster>",
		`<p:hf dt="0" ftr="0" hdr="0"/></p:sldMaster>`, 1)
	f.Pkg.Store(slideMasterTestPath, []byte(withHF))

	if err := f.SetSlideMasterBackground(slideMasterTestID, DecodeBackground{
		BGProperties: &DecodeBGProperties{
			SolidFill: &DecodeSolidFill{SolidRGBColor: &SolidRGBColor{Val: "0563C1"}},
		},
	}); err != nil {
		t.Fatal(err)
	}
	out := filepath.Join(t.TempDir(), "master_hf.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	master := string(readZipPart(t, out, slideMasterTestPath))
	if !strings.Contains(master, `<p:hf dt="0" ftr="0" hdr="0"/>`) {
		t.Fatalf("the header footer element is not kept on the round trip:\n%s", master)
	}
}

// TestMasterTransitionGuard checks that the modification of the slide master
// with the transition element returns the ErrSlideMasterUnsupported error
// (the attributes of the transition element are not modeled by the master
// round trip, the same as in the slide layout round trip).
func TestMasterTransitionGuard(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// the template master has no transition element: it is injected
	orig := string(f.readBytes(slideMasterTestPath))
	withTransition := strings.Replace(orig, "</p:sldMaster>",
		`<p:transition spd="med"/></p:sldMaster>`, 1)
	f.Pkg.Store(slideMasterTestPath, []byte(withTransition))

	if _, err := f.CreateMasterShape(slideMasterTestID, DecodeShapeProperties{}, DecodeTextBody{}); !errors.As(err, new(ErrSlideMasterUnsupported)) {
		t.Fatalf("expected ErrSlideMasterUnsupported, got %v", err)
	}
	if err := f.SetSlideMasterBackground(slideMasterTestID, DecodeBackground{}); !errors.As(err, new(ErrSlideMasterUnsupported)) {
		t.Fatalf("expected ErrSlideMasterUnsupported, got %v", err)
	}
}

// TestMasterShapeOrder checks that GetMasterShapeOrder reads the z-order of
// the shapes of the slide master, MoveMasterShape moves the shape before the
// target shape and SetMasterShapeOrder sets the order of the shapes.
func TestMasterShapeOrder(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// the ids of the shapes of the template master are 0–4 (in the document
	// order)
	shapeOrder, err := f.GetMasterShapeOrder(slideMasterTestID)
	if err != nil {
		t.Fatal(err)
	}
	wantIDs := []int{0, 1, 2, 3, 4}
	if len(shapeOrder) != len(wantIDs) {
		t.Fatalf("GetMasterShapeOrder = %v, want %d entries", shapeOrder, len(wantIDs))
	}
	for i, entry := range shapeOrder {
		if entry.Kind != ShapeOrderSp || entry.ID != wantIDs[i] {
			t.Fatalf("GetMasterShapeOrder[%d] = %+v, want sp %d", i, entry, wantIDs[i])
		}
	}

	// the new shape is appended to the end of the shape tree (the id 5)
	shapeID, err := f.CreateMasterShape(slideMasterTestID, DecodeShapeProperties{
		Xfrm: &DecodeXfrm{
			Offset:  &Offset{X: 0, Y: 0},
			Extents: &Extents{CX: 914400, CY: 457200},
		},
	}, DecodeTextBody{Paragraph: []DecodeParagraph{{Runs: []DecodeRuns{{Text: "The moved master shape"}}}}})
	if err != nil {
		t.Fatal(err)
	}
	if shapeID != 5 {
		t.Fatalf("CreateMasterShape id = %d, want 5", shapeID)
	}

	// the new shape is moved before the first shape of the master
	if err := f.MoveMasterShape(slideMasterTestID, shapeID, wantIDs[0]); err != nil {
		t.Fatal(err)
	}
	shapeOrder, err = f.GetMasterShapeOrder(slideMasterTestID)
	if err != nil {
		t.Fatal(err)
	}
	wantIDs = []int{5, 0, 1, 2, 3, 4}
	if len(shapeOrder) != len(wantIDs) {
		t.Fatalf("GetMasterShapeOrder after the move = %v, want %v", shapeOrder, wantIDs)
	}
	for i, entry := range shapeOrder {
		if entry.Kind != ShapeOrderSp || entry.ID != wantIDs[i] {
			t.Fatalf("GetMasterShapeOrder after the move[%d] = %+v, want sp %d", i, entry, wantIDs[i])
		}
	}

	out := filepath.Join(t.TempDir(), "master_order.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}

	// the order of the shapes is kept in the saved master (the id of the
	// first shape of the saved shape tree is the id of the moved shape)
	f2, err := OpenFile(out)
	if err != nil {
		t.Fatal(err)
	}
	defer f2.Close()
	shapeOrder, err = f2.GetMasterShapeOrder(slideMasterTestID)
	if err != nil {
		t.Fatal(err)
	}
	for i, entry := range shapeOrder {
		if entry.Kind != ShapeOrderSp || entry.ID != wantIDs[i] {
			t.Fatalf("the order of the shapes of the master is not kept: %v, want %v", shapeOrder, wantIDs)
		}
	}

	// the invalid order of the shapes (the entry of the shape tree is
	// missing)
	invalid := append([]ShapeOrder{}, shapeOrder[:len(shapeOrder)-1]...)
	if err := f.SetMasterShapeOrder(slideMasterTestID, invalid); !errors.Is(err, ErrShapeOrderInvalid) {
		t.Fatalf("expected ErrShapeOrderInvalid, got %v", err)
	}
	// the moved shape that is not in the shape tree
	if err := f.MoveMasterShape(slideMasterTestID, 100, wantIDs[0]); !errors.As(err, new(ErrShapeNotExist)) {
		t.Fatalf("expected ErrShapeNotExist, got %v", err)
	}
	// the non existing master
	if _, err := f.GetMasterShapeOrder(1); !errors.As(err, new(ErrSlideMasterNotExist)) {
		t.Fatalf("expected ErrSlideMasterNotExist, got %v", err)
	}
}

// TestLayoutShapeOrder checks that GetLayoutShapeOrder reads the z-order of
// the shapes of the slide layout (the shape tree of the template layout has
// the duplicated shape ids) and MoveLayoutShape / SetLayoutShapeOrder work
// with the shape tree with the duplicated shape ids.
func TestLayoutShapeOrder(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// the ids of the shapes of the template layout are 5, 6, 4, 5, 6 (in the
	// document order): the title and the slide number placeholder share the
	// id 5, the subtitle and the date placeholder share the id 6
	shapeOrder, err := f.GetLayoutShapeOrder(slideLayoutTestID)
	if err != nil {
		t.Fatal(err)
	}
	wantIDs := []int{5, 6, 4, 5, 6}
	if len(shapeOrder) != len(wantIDs) {
		t.Fatalf("GetLayoutShapeOrder = %v, want %d entries", shapeOrder, len(wantIDs))
	}
	for i, entry := range shapeOrder {
		if entry.Kind != ShapeOrderSp || entry.ID != wantIDs[i] {
			t.Fatalf("GetLayoutShapeOrder[%d] = %+v, want sp %d", i, entry, wantIDs[i])
		}
	}

	// the identity order (the same order the shape tree has) is set: the
	// shape tree with the duplicated shape ids is reordered to the same
	// order (the multiset check of the given order)
	if err := f.SetLayoutShapeOrder(slideLayoutTestID, shapeOrder); err != nil {
		t.Fatal(err)
	}

	// the new shape is appended to the end of the shape tree (the id 7) and
	// is moved before the shape with the id 5 (the last entry with the id 5:
	// the slide number placeholder)
	shapeID, err := f.CreateLayoutShape(slideLayoutTestID, DecodeShapeProperties{
		Xfrm: &DecodeXfrm{
			Offset:  &Offset{X: 0, Y: 0},
			Extents: &Extents{CX: 914400, CY: 457200},
		},
	}, DecodeTextBody{Paragraph: []DecodeParagraph{{Runs: []DecodeRuns{{Text: "The moved layout shape"}}}}})
	if err != nil {
		t.Fatal(err)
	}
	if shapeID != 7 {
		t.Fatalf("CreateLayoutShape id = %d, want 7", shapeID)
	}
	if err := f.MoveLayoutShape(slideLayoutTestID, shapeID, 5); err != nil {
		t.Fatal(err)
	}
	shapeOrder, err = f.GetLayoutShapeOrder(slideLayoutTestID)
	if err != nil {
		t.Fatal(err)
	}
	wantIDs = []int{5, 6, 4, 7, 5, 6}
	if len(shapeOrder) != len(wantIDs) {
		t.Fatalf("GetLayoutShapeOrder after the move = %v, want %v", shapeOrder, wantIDs)
	}
	for i, entry := range shapeOrder {
		if entry.Kind != ShapeOrderSp || entry.ID != wantIDs[i] {
			t.Fatalf("GetLayoutShapeOrder after the move[%d] = %+v, want sp %d", i, entry, wantIDs[i])
		}
	}

	out := filepath.Join(t.TempDir(), "layout_order.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}

	// the order of the shapes is kept in the saved layout
	f2, err := OpenFile(out)
	if err != nil {
		t.Fatal(err)
	}
	defer f2.Close()
	shapeOrder, err = f2.GetLayoutShapeOrder(slideLayoutTestID)
	if err != nil {
		t.Fatal(err)
	}
	for i, entry := range shapeOrder {
		if entry.Kind != ShapeOrderSp || entry.ID != wantIDs[i] {
			t.Fatalf("the order of the shapes of the layout is not kept: %v, want %v", shapeOrder, wantIDs)
		}
	}

	// the invalid order of the shapes (the duplicated entry of the shape
	// tree)
	invalid := append([]ShapeOrder{}, shapeOrder...)
	invalid = append(invalid, ShapeOrder{Kind: ShapeOrderSp, ID: 6})
	if err := f.SetLayoutShapeOrder(slideLayoutTestID, invalid); !errors.Is(err, ErrShapeOrderInvalid) {
		t.Fatalf("expected ErrShapeOrderInvalid, got %v", err)
	}
	// the moved shape that is not in the shape tree
	if err := f.MoveLayoutShape(slideLayoutTestID, 100, 5); !errors.As(err, new(ErrShapeNotExist)) {
		t.Fatalf("expected ErrShapeNotExist, got %v", err)
	}
	// the non existing layout
	if _, err := f.GetLayoutShapeOrder(1); !errors.As(err, new(ErrSlideLayoutNotExist)) {
		t.Fatalf("expected ErrSlideLayoutNotExist, got %v", err)
	}
}

// TestMasterShapeHyperLink checks that SetMasterShapeHyperLink sets the
// external and the internal hyperlink to the shape of the slide master,
// GetMasterShapeHyperLink / GetMasterShapeHyperLinks read the hyperlink back
// and LinkTypeNone removes the hyperlink (the relationship of the hyperlink
// is added to the rels part of the slide master and is removed with the
// hyperlink).
func TestMasterShapeHyperLink(t *testing.T) {
	f := NewFile()
	defer f.Close()

	shapeID, err := f.CreateMasterShape(slideMasterTestID, DecodeShapeProperties{
		Xfrm: &DecodeXfrm{
			Offset:  &Offset{X: 457200, Y: 457200},
			Extents: &Extents{CX: 914400, CY: 457200},
		},
	}, DecodeTextBody{Paragraph: []DecodeParagraph{{Runs: []DecodeRuns{{Text: "The master link"}}}}})
	if err != nil {
		t.Fatal(err)
	}

	// the external hyperlink (the hyperlink relationship with the
	// TargetMode "External" is added to the rels part of the slide master)
	tooltip := "gopptx on GitHub"
	const linkURL = "https://github.com/kenny-not-dead/gopptx"
	if err := f.SetMasterShapeHyperLink(slideMasterTestID, shapeID, linkURL, LinkTypeExternal,
		HyperlinkOpts{Tooltip: &tooltip}); err != nil {
		t.Fatal(err)
	}
	hasLink, link, err := f.GetMasterShapeHyperLink(slideMasterTestID, shapeID)
	if err != nil {
		t.Fatal(err)
	}
	if !hasLink || link != linkURL {
		t.Fatalf("GetMasterShapeHyperLink = %v %q, want true %q", hasLink, link, linkURL)
	}
	if ids, err := f.GetMasterShapeHyperLinks(slideMasterTestID, LinkTypeExternal); err != nil {
		t.Fatal(err)
	} else if len(ids) != 1 || ids[0] != shapeID {
		t.Fatalf("GetMasterShapeHyperLinks(External) = %v, want [%d]", ids, shapeID)
	}

	out := filepath.Join(t.TempDir(), "master_link_external.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	// the hyperlink relationship in the rels part of the slide master
	masterRels := string(readZipPart(t, out, slideMasterTestRelsPath))
	if !strings.Contains(masterRels, `relationships/hyperlink" Target="`+linkURL+`" TargetMode="External"`) {
		t.Fatalf("the hyperlink relationship is not added to the rels part of the slide master:\n%s", masterRels)
	}
	// the hyperlink element of the shape of the slide master
	master := string(readZipPart(t, out, slideMasterTestPath))
	if !strings.Contains(master, `<a:hlinkClick r:id="`) || !strings.Contains(master, `tooltip="gopptx on GitHub"`) {
		t.Fatalf("the hyperlink element is not written to the shape of the slide master:\n%s", master)
	}

	// the internal hyperlink to the slide of the presentation (the
	// relationship of the type "slide" without the TargetMode is added to
	// the rels part of the slide master, the target is the relative path of
	// the target slide part)
	const targetSlideID = 256
	if err := f.SetMasterShapeHyperLink(slideMasterTestID, shapeID, strconv.Itoa(targetSlideID), LinkTypeLocation); err != nil {
		t.Fatal(err)
	}
	hasLink, link, err = f.GetMasterShapeHyperLink(slideMasterTestID, shapeID)
	if err != nil {
		t.Fatal(err)
	}
	if !hasLink || link != strconv.Itoa(targetSlideID) {
		t.Fatalf("GetMasterShapeHyperLink = %v %q, want true %d", hasLink, link, targetSlideID)
	}
	if ids, err := f.GetMasterShapeHyperLinks(slideMasterTestID, LinkTypeLocation); err != nil {
		t.Fatal(err)
	} else if len(ids) != 1 || ids[0] != shapeID {
		t.Fatalf("GetMasterShapeHyperLinks(Location) = %v, want [%d]", ids, shapeID)
	}

	out = filepath.Join(t.TempDir(), "master_link_location.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	masterRels = string(readZipPart(t, out, slideMasterTestRelsPath))
	if !strings.Contains(masterRels, `relationships/slide" Target="../slides/slide1.xml"`) {
		t.Fatalf("the relationship to the target slide is not added to the rels part of the slide master:\n%s", masterRels)
	}
	master = string(readZipPart(t, out, slideMasterTestPath))
	if !strings.Contains(master, `action="ppaction://hlinksldjump"`) {
		t.Fatalf("the slide jump action is not written to the hyperlink element:\n%s", master)
	}

	// the hyperlink is kept on the round trip of the saved file
	f2, err := OpenFile(out)
	if err != nil {
		t.Fatal(err)
	}
	defer f2.Close()
	hasLink, link, err = f2.GetMasterShapeHyperLink(slideMasterTestID, shapeID)
	if err != nil {
		t.Fatal(err)
	}
	if !hasLink || link != strconv.Itoa(targetSlideID) {
		t.Fatalf("GetMasterShapeHyperLink after the round trip = %v %q, want true %d", hasLink, link, targetSlideID)
	}

	// the hyperlink is removed (the relationships of the hyperlink are
	// removed from the rels part of the slide master)
	if err := f.SetMasterShapeHyperLink(slideMasterTestID, shapeID, "", LinkTypeNone); err != nil {
		t.Fatal(err)
	}
	hasLink, link, err = f.GetMasterShapeHyperLink(slideMasterTestID, shapeID)
	if err != nil {
		t.Fatal(err)
	}
	if hasLink || link != "" {
		t.Fatalf("the hyperlink is not removed: %v %q", hasLink, link)
	}
	out = filepath.Join(t.TempDir(), "master_link_none.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	masterRels = string(readZipPart(t, out, slideMasterTestRelsPath))
	if strings.Contains(masterRels, "relationships/hyperlink") ||
		strings.Contains(masterRels, `relationships/slide" Target="../slides`) {
		t.Fatalf("the relationships of the hyperlink are not removed from the rels part of the slide master:\n%s", masterRels)
	}
	// the relationships of the slide layouts of the master are kept
	if !strings.Contains(masterRels, `relationships/slideLayout"`) {
		t.Fatalf("the relationships of the slide layouts are removed from the rels part of the slide master:\n%s", masterRels)
	}

	// the non existing master and the shape
	if _, _, err := f.GetMasterShapeHyperLink(1, shapeID); !errors.As(err, new(ErrSlideMasterNotExist)) {
		t.Fatalf("expected ErrSlideMasterNotExist, got %v", err)
	}
	if _, _, err := f.GetMasterShapeHyperLink(slideMasterTestID, 100); !errors.As(err, new(ErrShapeNotExist)) {
		t.Fatalf("expected ErrShapeNotExist, got %v", err)
	}
}

// TestLayoutShapeHyperLink checks that SetLayoutShapeHyperLink sets the
// external and the internal hyperlink to the shape of the slide layout,
// GetLayoutShapeHyperLink / GetLayoutShapeHyperLinks read the hyperlink back
// and LinkTypeNone removes the hyperlink (the relationship of the hyperlink
// is added to the rels part of the slide layout and is removed with the
// hyperlink).
func TestLayoutShapeHyperLink(t *testing.T) {
	f := NewFile()
	defer f.Close()

	shapeID, err := f.CreateLayoutShape(slideLayoutTestID, DecodeShapeProperties{
		Xfrm: &DecodeXfrm{
			Offset:  &Offset{X: 457200, Y: 457200},
			Extents: &Extents{CX: 914400, CY: 457200},
		},
	}, DecodeTextBody{Paragraph: []DecodeParagraph{{Runs: []DecodeRuns{{Text: "The layout link"}}}}})
	if err != nil {
		t.Fatal(err)
	}

	// the external hyperlink (the hyperlink relationship with the
	// TargetMode "External" is added to the rels part of the slide layout)
	tooltip := "gopptx on GitHub"
	const linkURL = "https://github.com/kenny-not-dead/gopptx"
	if err := f.SetLayoutShapeHyperLink(slideLayoutTestID, shapeID, linkURL, LinkTypeExternal,
		HyperlinkOpts{Tooltip: &tooltip}); err != nil {
		t.Fatal(err)
	}
	hasLink, link, err := f.GetLayoutShapeHyperLink(slideLayoutTestID, shapeID)
	if err != nil {
		t.Fatal(err)
	}
	if !hasLink || link != linkURL {
		t.Fatalf("GetLayoutShapeHyperLink = %v %q, want true %q", hasLink, link, linkURL)
	}
	if ids, err := f.GetLayoutShapeHyperLinks(slideLayoutTestID, LinkTypeExternal); err != nil {
		t.Fatal(err)
	} else if len(ids) != 1 || ids[0] != shapeID {
		t.Fatalf("GetLayoutShapeHyperLinks(External) = %v, want [%d]", ids, shapeID)
	}

	out := filepath.Join(t.TempDir(), "layout_link_external.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	// the hyperlink relationship in the rels part of the slide layout
	layoutRels := string(readZipPart(t, out, "ppt/slideLayouts/_rels/slideLayout1.xml.rels"))
	if !strings.Contains(layoutRels, `relationships/hyperlink" Target="`+linkURL+`" TargetMode="External"`) {
		t.Fatalf("the hyperlink relationship is not added to the rels part of the slide layout:\n%s", layoutRels)
	}
	// the hyperlink element of the shape of the slide layout
	layout := string(readZipPart(t, out, layoutTestPath))
	if !strings.Contains(layout, `<a:hlinkClick r:id="`) || !strings.Contains(layout, `tooltip="gopptx on GitHub"`) {
		t.Fatalf("the hyperlink element is not written to the shape of the slide layout:\n%s", layout)
	}

	// the internal hyperlink to the slide of the presentation (the
	// relationship of the type "slide" without the TargetMode is added to
	// the rels part of the slide layout, the target is the relative path of
	// the target slide part)
	const targetSlideID = 256
	if err := f.SetLayoutShapeHyperLink(slideLayoutTestID, shapeID, strconv.Itoa(targetSlideID), LinkTypeLocation); err != nil {
		t.Fatal(err)
	}
	hasLink, link, err = f.GetLayoutShapeHyperLink(slideLayoutTestID, shapeID)
	if err != nil {
		t.Fatal(err)
	}
	if !hasLink || link != strconv.Itoa(targetSlideID) {
		t.Fatalf("GetLayoutShapeHyperLink = %v %q, want true %d", hasLink, link, targetSlideID)
	}
	if ids, err := f.GetLayoutShapeHyperLinks(slideLayoutTestID, LinkTypeLocation); err != nil {
		t.Fatal(err)
	} else if len(ids) != 1 || ids[0] != shapeID {
		t.Fatalf("GetLayoutShapeHyperLinks(Location) = %v, want [%d]", ids, shapeID)
	}

	out = filepath.Join(t.TempDir(), "layout_link_location.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	layoutRels = string(readZipPart(t, out, "ppt/slideLayouts/_rels/slideLayout1.xml.rels"))
	if !strings.Contains(layoutRels, `relationships/slide" Target="../slides/slide1.xml"`) {
		t.Fatalf("the relationship to the target slide is not added to the rels part of the slide layout:\n%s", layoutRels)
	}
	layout = string(readZipPart(t, out, layoutTestPath))
	if !strings.Contains(layout, `action="ppaction://hlinksldjump"`) {
		t.Fatalf("the slide jump action is not written to the hyperlink element:\n%s", layout)
	}

	// the hyperlink is kept on the round trip of the saved file
	f2, err := OpenFile(out)
	if err != nil {
		t.Fatal(err)
	}
	defer f2.Close()
	hasLink, link, err = f2.GetLayoutShapeHyperLink(slideLayoutTestID, shapeID)
	if err != nil {
		t.Fatal(err)
	}
	if !hasLink || link != strconv.Itoa(targetSlideID) {
		t.Fatalf("GetLayoutShapeHyperLink after the round trip = %v %q, want true %d", hasLink, link, targetSlideID)
	}

	// the hyperlink is removed (the relationships of the hyperlink are
	// removed from the rels part of the slide layout)
	if err := f.SetLayoutShapeHyperLink(slideLayoutTestID, shapeID, "", LinkTypeNone); err != nil {
		t.Fatal(err)
	}
	hasLink, link, err = f.GetLayoutShapeHyperLink(slideLayoutTestID, shapeID)
	if err != nil {
		t.Fatal(err)
	}
	if hasLink || link != "" {
		t.Fatalf("the hyperlink is not removed: %v %q", hasLink, link)
	}
	out = filepath.Join(t.TempDir(), "layout_link_none.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	layoutRels = string(readZipPart(t, out, "ppt/slideLayouts/_rels/slideLayout1.xml.rels"))
	if strings.Contains(layoutRels, "relationships/hyperlink") ||
		strings.Contains(layoutRels, `relationships/slide" Target="../slides`) {
		t.Fatalf("the relationships of the hyperlink are not removed from the rels part of the slide layout:\n%s", layoutRels)
	}
	// the relationship of the slide master of the layout is kept
	if !strings.Contains(layoutRels, `relationships/slideMaster"`) {
		t.Fatalf("the relationship of the slide master is removed from the rels part of the slide layout:\n%s", layoutRels)
	}

	// the non existing layout and the shape
	if _, _, err := f.GetLayoutShapeHyperLink(1, shapeID); !errors.As(err, new(ErrSlideLayoutNotExist)) {
		t.Fatalf("expected ErrSlideLayoutNotExist, got %v", err)
	}
	if _, _, err := f.GetLayoutShapeHyperLink(slideLayoutTestID, 100); !errors.As(err, new(ErrShapeNotExist)) {
		t.Fatalf("expected ErrShapeNotExist, got %v", err)
	}
}

// TestNewSlideMaster checks that the NewSlideMaster function creates the new
// slide master as the copy of the source master (the "duplicate master" of
// PowerPoint): the new master part is the byte copy of the source master
// part with the ids of the slide layouts of the new master rewritten, the
// new layouts are the byte copies of the layouts of the source master
// pointing to the new master, and the content types, the rels parts, the
// slide master id list of the presentation and the rels of the presentation
// are updated. The file is stable on the round trip.
func TestNewSlideMaster(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// the hyperlink of a shape of the source master (the hyperlink
	// relationship is added to the rels part of the source master and must
	// be copied with the new master)
	const linkURL = "https://github.com/kenny-not-dead/gopptx"
	if err := f.SetMasterShapeHyperLink(slideMasterTestID, 0, linkURL, LinkTypeExternal); err != nil {
		t.Fatal(err)
	}

	newMasterID, err := f.NewSlideMaster(slideMasterTestID)
	if err != nil {
		t.Fatal(err)
	}
	if newMasterID != slideMasterTestID+2 {
		t.Fatalf("NewSlideMaster id = %d, want %d", newMasterID, slideMasterTestID+2)
	}

	out := filepath.Join(t.TempDir(), "new_master.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	// the part of the new master: the copy of the part of the source master
	// (the id of the slide layout of the new master in the slide layout id
	// list is rewritten, the rId is kept)
	newMasterPart := "ppt/slideMasters/slideMaster2.xml"
	wantMaster := bytes.ReplaceAll(readZipPart(t, out, slideMasterTestPath),
		[]byte(`<p:sldLayoutId id="`+strconv.Itoa(slideLayoutTestID)+`"`),
		[]byte(`<p:sldLayoutId id="`+strconv.Itoa(slideLayoutTestID+2)+`"`))
	if !bytes.Equal(readZipPart(t, out, newMasterPart), wantMaster) {
		t.Fatalf("the part of the new master is not the copy of the source master with the rewritten layout id")
	}
	// the part of the new layout: the byte copy of the part of the source
	// layout (the rels part of the new layout points to the new master)
	newLayoutPart := "ppt/slideLayouts/slideLayout2.xml"
	if !bytes.Equal(readZipPart(t, out, newLayoutPart), readZipPart(t, out, layoutTestPath)) {
		t.Fatalf("the part of the new layout is not the copy of the source layout")
	}
	newLayoutRels := string(readZipPart(t, out, "ppt/slideLayouts/_rels/slideLayout2.xml.rels"))
	if !strings.Contains(newLayoutRels, `relationships/slideMaster" Target="../slideMasters/slideMaster2.xml"`) {
		t.Fatalf("the rels part of the new layout is not created:\n%s", newLayoutRels)
	}
	// the rels part of the new master: the copy of the rels part of the
	// source master (the target of the relationship of the slide layout is
	// rewritten to the new layout)
	newMasterRels := string(readZipPart(t, out, "ppt/slideMasters/_rels/slideMaster2.xml.rels"))
	if !strings.Contains(newMasterRels, `relationships/slideLayout" Target="../slideLayouts/slideLayout2.xml"`) {
		t.Fatalf("the rels part of the new master is not created:\n%s", newMasterRels)
	}
	if !strings.Contains(newMasterRels, `relationships/theme"`) {
		t.Fatalf("the theme relationship is not kept in the rels part of the new master:\n%s", newMasterRels)
	}
	// the hyperlink relationship of the source master is copied with the
	// new master
	if !strings.Contains(newMasterRels, `relationships/hyperlink" Target="`+linkURL+`" TargetMode="External"`) {
		t.Fatalf("the hyperlink relationship is not copied to the rels part of the new master:\n%s", newMasterRels)
	}
	// the content types of the new master, its rels part, the new layout and
	// its rels part
	contentTypes := string(readZipPart(t, out, "[Content_Types].xml"))
	if !strings.Contains(contentTypes, `ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml" PartName="/ppt/slideMasters/slideMaster2.xml"`) ||
		!strings.Contains(contentTypes, `ContentType="application/vnd.openxmlformats-package.relationships+xml" PartName="/ppt/slideMasters/_rels/slideMaster2.xml.rels"`) ||
		!strings.Contains(contentTypes, `ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml" PartName="/ppt/slideLayouts/slideLayout2.xml"`) ||
		!strings.Contains(contentTypes, `ContentType="application/vnd.openxmlformats-package.relationships+xml" PartName="/ppt/slideLayouts/_rels/slideLayout2.xml.rels"`) {
		t.Fatalf("the content types of the new master and the new layout are not added:\n%s", contentTypes)
	}
	// the entry of the new master in the slide master id list of the
	// presentation (the relationship of the new master in the rels of the
	// presentation)
	presentation := string(readZipPart(t, out, "ppt/presentation.xml"))
	if !strings.Contains(presentation, `<p:sldMasterId id="`+strconv.Itoa(slideMasterTestID)+`" r:id="rId2"/>`) ||
		!strings.Contains(presentation, `<p:sldMasterId id="`+strconv.Itoa(newMasterID)+`" r:id="rId5"/>`) {
		t.Fatalf("the entry of the new master is not added to the slide master id list of the presentation:\n%s", presentation)
	}
	presentationRels := string(readZipPart(t, out, "ppt/_rels/presentation.xml.rels"))
	if !strings.Contains(presentationRels, `relationships/slideMaster" Target="slideMasters/slideMaster2.xml"`) {
		t.Fatalf("the relationship of the new master is not added to the rels of the presentation:\n%s", presentationRels)
	}
	// the lists of the masters and the layouts of the presentation
	if list := f.GetSlideMasterList(); len(list) != 2 || list[1] != newMasterID {
		t.Fatalf("GetSlideMasterList = %v, want [%d %d]", list, slideMasterTestID, newMasterID)
	}
	if list := f.GetSlideLayoutList(); len(list) != 2 || list[1] != slideLayoutTestID+2 {
		t.Fatalf("GetSlideLayoutList = %v, want [%d %d]", list, slideLayoutTestID, slideLayoutTestID+2)
	}

	// the second new master (the id is incremented again, the part is
	// created with the next free index)
	secondMasterID, err := f.NewSlideMaster(newMasterID)
	if err != nil {
		t.Fatal(err)
	}
	if secondMasterID != slideMasterTestID+4 {
		t.Fatalf("the second NewSlideMaster id = %d, want %d", secondMasterID, slideMasterTestID+4)
	}
	if list := f.GetSlideLayoutList(); len(list) != 3 || list[2] != slideLayoutTestID+4 {
		t.Fatalf("GetSlideLayoutList after the second NewSlideMaster = %v, want 3 entries", list)
	}

	// the file is stable on the round trip: the lists of the masters and
	// the layouts are kept, the parts of the new master and the new layout
	// are kept
	f2, err := OpenFile(out)
	if err != nil {
		t.Fatal(err)
	}
	defer f2.Close()
	if list := f2.GetSlideMasterList(); len(list) != 2 || list[1] != newMasterID {
		t.Fatalf("GetSlideMasterList after the round trip = %v, want [%d %d]", list, slideMasterTestID, newMasterID)
	}
	if list := f2.GetSlideLayoutList(); len(list) != 2 || list[1] != slideLayoutTestID+2 {
		t.Fatalf("GetSlideLayoutList after the round trip = %v, want [%d %d]", list, slideLayoutTestID, slideLayoutTestID+2)
	}
	out2 := filepath.Join(t.TempDir(), "new_master_roundtrip.pptx")
	if err := f2.SaveAs(out2); err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(readZipPart(t, out2, newMasterPart), readZipPart(t, out, newMasterPart)) ||
		!bytes.Equal(readZipPart(t, out2, newLayoutPart), readZipPart(t, out, newLayoutPart)) {
		t.Fatalf("the parts of the new master and the new layout are not kept on the round trip")
	}

	// the non existing source master
	if _, err := f.NewSlideMaster(1); !errors.As(err, new(ErrSlideMasterNotExist)) {
		t.Fatalf("expected ErrSlideMasterNotExist, got %v", err)
	}
}
