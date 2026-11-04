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
