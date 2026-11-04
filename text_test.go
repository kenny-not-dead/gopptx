// Copyright 2026 kenny-not-dead. All rights reserved.
// Use of this source code is governed by a BSD-style license
// that can be found in the LICENSE file.
//
// Package gopptx provides functionality to create and manipulate PowerPoint
// (.pptx) files in Go, using the Office Open XML (ECMA-376) format.

package gopptx

import (
	"errors"
	"path/filepath"
	"regexp"
	"strings"
	"testing"
)

// titleShapeID provides a function to get the id of the title placeholder
// (p:ph type="title" or "ctrTitle") of the slide.
func titleShapeID(t *testing.T, f *File, slideID int) int {
	t.Helper()
	shapes, err := f.GetShapes(slideID)
	if err != nil {
		t.Fatal(err)
	}
	for _, shape := range shapes {
		if shape.NonVisualShapeProperties == nil ||
			shape.NonVisualShapeProperties.NonVisualProperties == nil ||
			shape.NonVisualShapeProperties.NonVisualProperties.Ph == nil ||
			shape.NonVisualShapeProperties.NonVisualProperties.Ph.Type == nil {
			continue
		}
		switch ph := *shape.NonVisualShapeProperties.NonVisualProperties.Ph.Type; ph {
		case "title", "ctrTitle":
			return shape.NonVisualShapeProperties.CommonNonVisualProperties.ID
		}
	}
	t.Fatal("the title placeholder was not found")
	return 0
}

var guidRe = regexp.MustCompile(`^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$`)

// TestSetShapeText checks that SetShapeText sets the text of the shape as one
// run per line with the line break (a:br) between the lines, and that the
// text is preserved on save and on the next load.
func TestSetShapeText(t *testing.T) {
	f := NewFile()
	defer f.Close()
	slideID := f.GetSlideList()[0]
	titleID := titleShapeID(t, f, slideID)

	if err := f.SetShapeText(slideID, 999, "text"); !errors.As(err, new(ErrShapeNotExist)) {
		t.Fatalf("expected ErrShapeNotExist, got %v", err)
	}

	if err := f.SetShapeText(slideID, titleID, "Line 1\nLine 2"); err != nil {
		t.Fatal(err)
	}

	out := filepath.Join(t.TempDir(), "text.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	slideXML := string(readZipPart(t, out, "ppt/slides/slide1.xml"))
	// the runs and the line break are interleaved in the document order:
	// the run, the line break, the run
	if !strings.Contains(slideXML, "Line 1</a:t></a:r><a:br/><a:r><a:t>Line 2") {
		t.Fatalf("the line break is not placed between the runs in the slide XML:\n%s", slideXML)
	}

	f2, err := OpenFile(out)
	if err != nil {
		t.Fatal(err)
	}
	defer f2.Close()

	shapes, err := f2.GetShapes(slideID)
	if err != nil {
		t.Fatal(err)
	}
	var para *DecodeParagraph
	for i := range shapes {
		body := shapes[i].TextBody
		if body == nil || shapes[i].NonVisualShapeProperties == nil ||
			shapes[i].NonVisualShapeProperties.CommonNonVisualProperties == nil ||
			shapes[i].NonVisualShapeProperties.CommonNonVisualProperties.ID != titleID {
			continue
		}
		if len(body.Paragraph) > 0 {
			para = &body.Paragraph[0]
		}
	}
	if para == nil {
		t.Fatal("the title text body was not read back")
	}
	if len(para.Runs) != 2 || para.Runs[0].Text != "Line 1" || para.Runs[1].Text != "Line 2" {
		t.Fatalf("the runs of the paragraph are not preserved: %+v", para.Runs)
	}
	if len(para.Breaks) != 1 {
		t.Fatalf("expected 1 line break, got %d", len(para.Breaks))
	}
	// the document order of the paragraph children is preserved on load
	want := []paragraphOrderEntry{
		{kind: paragraphOrderRun, index: 0},
		{kind: paragraphOrderBr, index: 0},
		{kind: paragraphOrderRun, index: 1},
	}
	if len(para.paragraphOrder) != len(want) {
		t.Fatalf("the paragraph order = %v, want %v", para.paragraphOrder, want)
	}
	for i, entry := range want {
		if para.paragraphOrder[i] != entry {
			t.Fatalf("the paragraph order = %v, want %v", para.paragraphOrder, want)
		}
	}
}

// TestSlideNumberField checks that the slide number field (a:fld of the type
// "slidenum") of the paragraph is written on save and is preserved on the
// next load.
func TestSlideNumberField(t *testing.T) {
	f := NewFile()
	defer f.Close()
	slideID := f.GetSlideList()[0]

	field := NewSlideNumberField()
	if field.Type != FieldSlideNumber {
		t.Fatalf("the field type = %q, want %q", field.Type, FieldSlideNumber)
	}
	if !guidRe.MatchString(field.ID) {
		t.Fatalf("the field id %q is not a GUID", field.ID)
	}

	shapes, err := f.GetShapes(slideID)
	if err != nil {
		t.Fatal(err)
	}
	subtitleID := 0
	for _, shape := range shapes {
		if shape.NonVisualShapeProperties == nil ||
			shape.NonVisualShapeProperties.NonVisualProperties == nil ||
			shape.NonVisualShapeProperties.NonVisualProperties.Ph == nil ||
			shape.NonVisualShapeProperties.NonVisualProperties.Ph.Type == nil {
			continue
		}
		if ph := *shape.NonVisualShapeProperties.NonVisualProperties.Ph.Type; ph == "subTitle" {
			subtitleID = shape.NonVisualShapeProperties.CommonNonVisualProperties.ID
		}
	}
	if subtitleID == 0 {
		t.Fatal("the subtitle placeholder was not found")
	}

	if err := f.SetShapeTextBody(slideID, subtitleID, DecodeTextBody{
		BodyProperties: &DecodeBodyProperties{},
		Paragraph: []DecodeParagraph{{
			Field: []DecodeField{field},
		}},
	}); err != nil {
		t.Fatal(err)
	}

	out := filepath.Join(t.TempDir(), "slidenum.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	slideXML := string(readZipPart(t, out, "ppt/slides/slide1.xml"))
	if !strings.Contains(slideXML, `type="slidenum"`) {
		t.Fatalf("the slide number field is not written in the slide XML:\n%s", slideXML)
	}
	if !strings.Contains(slideXML, "id=\""+field.ID+"\"") {
		t.Fatalf("the id of the slide number field is not written in the slide XML:\n%s", slideXML)
	}

	f2, err := OpenFile(out)
	if err != nil {
		t.Fatal(err)
	}
	defer f2.Close()

	shapes, err = f2.GetShapes(slideID)
	if err != nil {
		t.Fatal(err)
	}
	for i := range shapes {
		body := shapes[i].TextBody
		if body == nil {
			continue
		}
		for j := range body.Paragraph {
			for _, fld := range body.Paragraph[j].Field {
				if fld.Type != FieldSlideNumber || fld.ID != field.ID || fld.Text != "1" {
					t.Fatalf("the slide number field is not preserved: %+v", fld)
				}
			}
		}
	}
}

// TestSetShapeTextBodyBuAutoNum checks that the auto numbering of the
// paragraph (a:buAutoNum) is written on save and is preserved on the next
// load.
func TestSetShapeTextBodyBuAutoNum(t *testing.T) {
	f := NewFile()
	defer f.Close()
	slideID := f.GetSlideList()[0]
	titleID := titleShapeID(t, f, slideID)

	startAt := 3
	if err := f.SetShapeTextBody(slideID, titleID, DecodeTextBody{
		BodyProperties: &DecodeBodyProperties{},
		Paragraph: []DecodeParagraph{
			{
				ParagraphProperties: &DecodeParagraphProperties{
					BuFont:    &DecodeBulletFont{Typeface: "Arial"},
					BuAutoNum: &DecodeBulletAutoNum{Type: "arabicPeriod"},
				},
				Runs: []DecodeRuns{{Text: "First item"}},
			},
			{
				ParagraphProperties: &DecodeParagraphProperties{
					BuFont:    &DecodeBulletFont{Typeface: "Arial"},
					BuAutoNum: &DecodeBulletAutoNum{Type: "arabicPeriod", StartAt: &startAt},
				},
				Runs: []DecodeRuns{{Text: "Second item"}},
			},
		},
	}); err != nil {
		t.Fatal(err)
	}

	out := filepath.Join(t.TempDir(), "autonum.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	slideXML := string(readZipPart(t, out, "ppt/slides/slide1.xml"))
	if !strings.Contains(slideXML, `<a:buAutoNum type="arabicPeriod"/>`) {
		t.Fatalf("the buAutoNum element is not written in the slide XML:\n%s", slideXML)
	}
	if !strings.Contains(slideXML, `<a:buAutoNum type="arabicPeriod" startAt="3"/>`) {
		t.Fatalf("the buAutoNum element with the startAt attribute is not written in the slide XML:\n%s", slideXML)
	}

	f2, err := OpenFile(out)
	if err != nil {
		t.Fatal(err)
	}
	defer f2.Close()

	shapes, err := f2.GetShapes(slideID)
	if err != nil {
		t.Fatal(err)
	}
	var body *DecodeTextBody
	for i := range shapes {
		shape := &shapes[i]
		if shape.NonVisualShapeProperties == nil ||
			shape.NonVisualShapeProperties.CommonNonVisualProperties == nil ||
			shape.NonVisualShapeProperties.CommonNonVisualProperties.ID != titleID {
			continue
		}
		body = shape.TextBody
	}
	if body == nil {
		t.Fatal("the title text body was not read back")
	}
	if len(body.Paragraph) != 2 {
		t.Fatalf("the number of paragraphs is %d, want 2", len(body.Paragraph))
	}
	for i, para := range body.Paragraph {
		if para.ParagraphProperties == nil || para.ParagraphProperties.BuAutoNum == nil ||
			para.ParagraphProperties.BuAutoNum.Type != "arabicPeriod" {
			t.Fatalf("the buAutoNum element of the paragraph %d is not preserved: %+v",
				i, para.ParagraphProperties)
		}
	}
	got := body.Paragraph[1].ParagraphProperties.BuAutoNum.StartAt
	if got == nil || *got != startAt {
		t.Fatalf("the startAt attribute is not preserved: %v", got)
	}
}

// TestLoadedTextBreakRoundTrip checks that the line break (a:br) and the
// slide number field (a:fld) of the real presentation (the fixture with the
// text box created by python-pptx) are preserved on the round trip: the
// paragraph is not modified, so the document order of its children is
// written back as it is in the origin slide part.
func TestLoadedTextBreakRoundTrip(t *testing.T) {
	fixture := filepath.Join("test", "textbreak.pptx")
	f, err := OpenFile(fixture)
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()

	// the line break and the field of the text box are read back
	shapes, err := f.GetShapes(256)
	if err != nil {
		t.Fatal(err)
	}
	var (
		foundBr  bool
		foundFld bool
	)
	for i := range shapes {
		body := shapes[i].TextBody
		if body == nil {
			continue
		}
		for j := range body.Paragraph {
			para := &body.Paragraph[j]
			if len(para.Breaks) == 1 && len(para.Runs) == 2 &&
				para.Runs[0].Text == "Line one" && para.Runs[1].Text == "Line two" {
				foundBr = true
				// the document order of the paragraph children: the run,
				// the line break, the run
				want := []paragraphOrderEntry{
					{kind: paragraphOrderRun, index: 0},
					{kind: paragraphOrderBr, index: 0},
					{kind: paragraphOrderRun, index: 1},
				}
				if len(para.paragraphOrder) != len(want) {
					t.Fatalf("the paragraph order = %v, want %v", para.paragraphOrder, want)
				}
				for k, entry := range want {
					if para.paragraphOrder[k] != entry {
						t.Fatalf("the paragraph order = %v, want %v", para.paragraphOrder, want)
					}
				}
			}
			for _, fld := range para.Field {
				if fld.Type == FieldSlideNumber {
					foundFld = true
					if fld.Text != "1" || !guidRe.MatchString(fld.ID) {
						t.Fatalf("the slide number field is not read back: %+v", fld)
					}
				}
			}
		}
	}
	if !foundBr {
		t.Fatal("the line break of the fixture is not read back")
	}
	if !foundFld {
		t.Fatal("the slide number field of the fixture is not read back")
	}

	out := filepath.Join(t.TempDir(), "textbreak.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}

	// the slide part with the line break and the field must be byte for byte
	// identical
	original := readZipPart(t, fixture, "ppt/slides/slide1.xml")
	saved := readZipPart(t, out, "ppt/slides/slide1.xml")
	if string(original) != string(saved) {
		t.Fatal("the slide part with the line break and the field is not preserved byte for byte on the round trip")
	}
}
