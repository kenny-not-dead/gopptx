// Copyright 2026 kenny-not-dead. All rights reserved.
// Use of this source code is governed by a BSD-style license
// that can be found in the LICENSE file.
//
// Package gopptx provides functionality to create and manipulate PowerPoint
// (.pptx) files in Go, using the Office Open XML (ECMA-376) format.

package gopptx

import (
	"encoding/xml"
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

// TestRunPropertiesItalic checks the italic attribute (a:rPr i="1") of the
// run properties: the attribute is kept by the conversion of the run
// properties, is exported to the slide part and to the notes slide part on
// the save and is re-imported from the saved file.
func TestRunPropertiesItalic(t *testing.T) {
	one := 1

	// the conversion of the run properties keeps the italic attribute
	rp := newRunProperties(&DecodeRunProperties{Italic: &one, Bold: &one, Lang: "en-US"})
	if rp == nil || rp.Italic == nil || *rp.Italic != 1 || rp.Bold == nil || *rp.Bold != 1 {
		t.Fatalf("newRunProperties did not keep the italic and the bold attributes: %+v", rp)
	}
	// the run properties without the italic attribute keep the attribute unset
	if rp2 := newRunProperties(&DecodeRunProperties{Bold: &one}); rp2.Italic != nil {
		t.Fatalf("newRunProperties set the italic attribute of the unset run properties: %+v", rp2)
	}

	f, err := OpenFile(testFixture)
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()

	// the italic run is set to the shape of the slide
	err = f.SetShapeTextBody(257, 8, DecodeTextBody{
		BodyProperties: &DecodeBodyProperties{},
		ListStyle:      &struct{}{},
		Paragraph: []DecodeParagraph{
			{Runs: []DecodeRuns{{Text: "italic", RunProperties: &DecodeRunProperties{Italic: &one}}}},
		},
	})
	if err != nil {
		t.Fatal(err)
	}

	// the italic run is set to the notes of the slide
	if err := f.SetNotes(257, DecodeTextBody{
		BodyProperties: &DecodeBodyProperties{},
		ListStyle:      &struct{}{},
		Paragraph: []DecodeParagraph{
			{Runs: []DecodeRuns{{Text: "italic notes", RunProperties: &DecodeRunProperties{Italic: &one}}}},
		},
	}); err != nil {
		t.Fatal(err)
	}

	out := filepath.Join(t.TempDir(), "italic.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}

	// the italic attribute is exported to the slide part
	slideXML := string(readZipPart(t, out, "ppt/slides/slide2.xml"))
	if !strings.Contains(slideXML, `i="1"`) {
		t.Errorf("the italic attribute is not exported to the slide part: %s", slideXML)
	}

	// the italic attribute is exported to the notes slide part (the notes
	// are re-serialized through the encode types on the save)
	notesPath, ok := f.getNotesSlidePath(257)
	if !ok {
		t.Fatal("the notes slide path of the slide was not found")
	}
	notesXML := string(readZipPart(t, out, notesPath))
	if !strings.Contains(notesXML, `i="1"`) {
		t.Errorf("the italic attribute is not exported to the notes slide part: %s", notesXML)
	}

	// the italic attribute is re-imported from the saved file
	f2, err := OpenFile(out)
	if err != nil {
		t.Fatal(err)
	}
	defer f2.Close()

	shapes, err := f2.GetShapes(257)
	if err != nil {
		t.Fatal(err)
	}
	for i := range shapes {
		if shapes[i].NonVisualShapeProperties.CommonNonVisualProperties.ID == 8 {
			rp2 := shapes[i].TextBody.Paragraph[0].Runs[0].RunProperties
			if rp2 == nil || rp2.Italic == nil || *rp2.Italic != 1 {
				t.Fatalf("the italic attribute is not re-imported from the saved slide: %+v", rp2)
			}
		}
	}

	notes, err := f2.GetNotes(257)
	if err != nil {
		t.Fatal(err)
	}
	rp3 := notes.Paragraph[0].Runs[0].RunProperties
	if rp3 == nil || rp3.Italic == nil || *rp3.Italic != 1 {
		t.Fatalf("the italic attribute is not re-imported from the saved notes: %+v", rp3)
	}
}

// TestParagraphTextRunsDocumentOrder checks that TextRuns returns the
// children of the paragraph (the runs, the line breaks and the fields) in
// the document order: the position of the line break (between which runs)
// is the index of the break in the returned slice, and the run properties
// (the formatting) of each element are returned with it.
func TestParagraphTextRunsDocumentOrder(t *testing.T) {
	const paragraphXML = `<a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
		`<a:r><a:rPr lang="en-US" b="1"/><a:t>Line 1 </a:t></a:r>` +
		`<a:br/>` +
		`<a:fld id="{00000000-0000-0000-0000-000000000001}" type="slidenum"><a:t>4</a:t></a:fld>` +
		`<a:r><a:t>Line 2</a:t></a:r>` +
		`</a:p>`

	var p DecodeParagraph
	if err := xml.Unmarshal([]byte(paragraphXML), &p); err != nil {
		t.Fatal(err)
	}

	runs := p.TextRuns()
	if len(runs) != 4 {
		t.Fatalf("the number of the text runs = %d, want 4: %+v", len(runs), runs)
	}

	wantKinds := []TextRunKind{TextRunText, TextRunBreak, TextRunField, TextRunText}
	for i, run := range runs {
		if run.Kind != wantKinds[i] {
			t.Fatalf("the kind of the text run %d = %d, want %d: %+v", i, run.Kind, wantKinds[i], runs)
		}
	}

	// the run keeps the text and the run properties of the run
	if runs[0].Run.Text != "Line 1 " {
		t.Fatalf("the text of the first run = %q, want %q", runs[0].Run.Text, "Line 1 ")
	}
	if runs[0].Run.RunProperties == nil ||
		runs[0].Run.RunProperties.Lang != "en-US" ||
		runs[0].Run.RunProperties.Bold == nil || *runs[0].Run.RunProperties.Bold != 1 {
		t.Fatalf("the run properties of the first run are not returned: %+v", runs[0].Run.RunProperties)
	}
	// the line break is the second text run (index 1): after the first run
	// and before the field
	if runs[1].Break.RunProperties != nil {
		t.Fatalf("the run properties of the line break are not empty: %+v", runs[1].Break.RunProperties)
	}
	// the field keeps the id, the type and the cached value of the field
	if runs[2].Field.Type != FieldSlideNumber ||
		runs[2].Field.ID != "{00000000-0000-0000-0000-000000000001}" ||
		runs[2].Field.Text != "4" {
		t.Fatalf("the field is not returned: %+v", runs[2].Field)
	}
	if runs[3].Run.Text != "Line 2" {
		t.Fatalf("the text of the last run = %q, want %q", runs[3].Run.Text, "Line 2")
	}
}

// TestParagraphTextRunsBlockOrder checks that TextRuns of the paragraph
// built by the caller (without the document order) returns the children as
// blocks: the runs, then the line breaks, then the fields (the same order
// the paragraph is serialized).
func TestParagraphTextRunsBlockOrder(t *testing.T) {
	p := DecodeParagraph{
		Runs:   []DecodeRuns{{Text: "first"}, {Text: "second"}},
		Breaks: []DecodeBr{{}},
		Field:  []DecodeField{{ID: "{00000000-0000-0000-0000-000000000001}", Type: FieldSlideNumber, Text: "1"}},
	}

	runs := p.TextRuns()
	if len(runs) != 4 {
		t.Fatalf("the number of the text runs = %d, want 4: %+v", len(runs), runs)
	}
	wantKinds := []TextRunKind{TextRunText, TextRunText, TextRunBreak, TextRunField}
	for i, run := range runs {
		if run.Kind != wantKinds[i] {
			t.Fatalf("the kind of the text run %d = %d, want %d: %+v", i, run.Kind, wantKinds[i], runs)
		}
	}
	if runs[0].Run.Text != "first" || runs[1].Run.Text != "second" {
		t.Fatalf("the text of the runs is not returned: %+v", runs)
	}
	if runs[3].Field.Type != FieldSlideNumber {
		t.Fatalf("the field is not returned: %+v", runs[3].Field)
	}
}

// TestParagraphTextRunsLoaded checks that TextRuns of the paragraph of the
// loaded presentation (the fixture with the text box created by
// python-pptx, the line break between the runs) returns the children of
// the paragraph in the document order of the slide part: the run, the line
// break (the index 1), the run.
func TestParagraphTextRunsLoaded(t *testing.T) {
	fixture := filepath.Join("test", "textbreak.pptx")
	f, err := OpenFile(fixture)
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()

	shapes, err := f.GetShapes(256)
	if err != nil {
		t.Fatal(err)
	}

	found := false
	for i := range shapes {
		body := shapes[i].TextBody
		if body == nil {
			continue
		}
		for j := range body.Paragraph {
			para := &body.Paragraph[j]
			if len(para.Runs) != 2 || len(para.Breaks) != 1 ||
				para.Runs[0].Text != "Line one" || para.Runs[1].Text != "Line two" {
				continue
			}

			runs := para.TextRuns()
			if len(runs) != 3 {
				t.Fatalf("the number of the text runs = %d, want 3: %+v", len(runs), runs)
			}
			wantKinds := []TextRunKind{TextRunText, TextRunBreak, TextRunText}
			for k, run := range runs {
				if run.Kind != wantKinds[k] {
					t.Fatalf("the kind of the text run %d = %d, want %d: %+v", k, run.Kind, wantKinds[k], runs)
				}
			}
			// the line break is between the runs: the text of the shape is
			// the runs with the line break (a:br) between them
			if runs[0].Run.Text != "Line one" || runs[2].Run.Text != "Line two" {
				t.Fatalf("the text of the runs is not preserved: %+v", runs)
			}
			found = true
		}
	}
	if !found {
		t.Fatal("the paragraph with the line break of the fixture was not found")
	}
}

// TestGetShapeText checks that GetShapeText returns the text of the shape
// as a flat string: the two shapes with the same visual text (two
// paragraphs, one paragraph with the line break) return the same string
// (the question of the issue), and the mixed text body (the runs, the empty
// paragraph and the line break with the field) is rendered as in the design
// of the issue.
func TestGetShapeText(t *testing.T) {
	f := NewFile()
	defer f.Close()
	slideID := f.GetSlideList()[0]
	titleID := titleShapeID(t, f, slideID)

	if _, err := f.GetShapeText(slideID, 999); !errors.As(err, new(ErrShapeNotExist)) {
		t.Fatalf("expected ErrShapeNotExist, got %v", err)
	}
	if _, err := f.GetShapeText(999, titleID); !errors.As(err, new(ErrSlideNotExist)) {
		t.Fatalf("expected ErrSlideNotExist, got %v", err)
	}

	// the issue example, slide 1: the text of two paragraphs
	if err := f.SetShapeTextBody(slideID, titleID, DecodeTextBody{
		BodyProperties: &DecodeBodyProperties{},
		Paragraph: []DecodeParagraph{
			{Runs: []DecodeRuns{{Text: "Line 1"}}},
			{Runs: []DecodeRuns{{Text: "Line 2"}}},
		},
	}); err != nil {
		t.Fatal(err)
	}
	text, err := f.GetShapeText(slideID, titleID)
	if err != nil {
		t.Fatal(err)
	}
	if text != "Line 1\nLine 2" {
		t.Fatalf("GetShapeText = %q, want %q", text, "Line 1\nLine 2")
	}

	// the issue example, slide 2: one paragraph with the line break
	if err := f.SetShapeText(slideID, titleID, "Line 1\nLine 2"); err != nil {
		t.Fatal(err)
	}
	text, err = f.GetShapeText(slideID, titleID)
	if err != nil {
		t.Fatal(err)
	}
	if text != "Line 1\nLine 2" {
		t.Fatalf("GetShapeText = %q, want %q", text, "Line 1\nLine 2")
	}

	// the mixed example of the issue: the runs of the paragraph, the empty
	// paragraph, the line break with the field
	if err := f.SetShapeTextBody(slideID, titleID, DecodeTextBody{
		BodyProperties: &DecodeBodyProperties{},
		Paragraph: []DecodeParagraph{
			{Runs: []DecodeRuns{{Text: "Foo "}, {Text: "Bar"}}},
			{},
			{
				Runs:   []DecodeRuns{{Text: "Line 1"}},
				Breaks: []DecodeBr{{}},
				Field:  []DecodeField{{ID: "{00000000-0000-0000-0000-000000000001}", Type: FieldSlideNumber, Text: "4"}},
			},
		},
	}); err != nil {
		t.Fatal(err)
	}
	text, err = f.GetShapeText(slideID, titleID)
	if err != nil {
		t.Fatal(err)
	}
	if text != "Foo Bar\n\nLine 1\n4" {
		t.Fatalf("GetShapeText = %q, want %q", text, "Foo Bar\n\nLine 1\n4")
	}
}

// TestGetShapeTextRuns checks that GetShapeTextRuns returns the paragraphs
// of the text body of the shape, each paragraph is the text runs in the
// document order of the paragraph: the position of the line break (between
// which runs) is the index of the break in the text runs of the paragraph.
// The shape without the text body returns an empty slice.
func TestGetShapeTextRuns(t *testing.T) {
	f := NewFile()
	defer f.Close()
	slideID := f.GetSlideList()[0]
	titleID := titleShapeID(t, f, slideID)

	if _, err := f.GetShapeTextRuns(slideID, 999); !errors.As(err, new(ErrShapeNotExist)) {
		t.Fatalf("expected ErrShapeNotExist, got %v", err)
	}
	if _, err := f.GetShapeTextRuns(999, titleID); !errors.As(err, new(ErrSlideNotExist)) {
		t.Fatalf("expected ErrSlideNotExist, got %v", err)
	}

	// the shape without the text body returns an empty slice, the text of
	// the shape is the empty string
	shapeID, err := f.CreateShape(slideID, DecodeShapeProperties{}, DecodeTextBody{})
	if err != nil {
		t.Fatal(err)
	}
	runs, err := f.GetShapeTextRuns(slideID, shapeID)
	if err != nil {
		t.Fatal(err)
	}
	if runs == nil || len(runs) != 0 {
		t.Fatalf("GetShapeTextRuns = %v, want an empty slice", runs)
	}
	if text, err := f.GetShapeText(slideID, shapeID); err != nil {
		t.Fatal(err)
	} else if text != "" {
		t.Fatalf("GetShapeText = %q, want an empty string", text)
	}

	// the text of the shape: one paragraph with the line break between the
	// runs
	if err := f.SetShapeText(slideID, titleID, "Line 1\nLine 2"); err != nil {
		t.Fatal(err)
	}
	runs, err = f.GetShapeTextRuns(slideID, titleID)
	if err != nil {
		t.Fatal(err)
	}
	if len(runs) != 1 {
		t.Fatalf("the number of the paragraphs = %d, want 1: %+v", len(runs), runs)
	}
	paragraph := runs[0]
	if len(paragraph) != 3 {
		t.Fatalf("the number of the text runs = %d, want 3: %+v", len(paragraph), paragraph)
	}
	wantKinds := []TextRunKind{TextRunText, TextRunBreak, TextRunText}
	for i, run := range paragraph {
		if run.Kind != wantKinds[i] {
			t.Fatalf("the kind of the text run %d = %d, want %d: %+v", i, run.Kind, wantKinds[i], paragraph)
		}
	}
	// the line break is the second text run (index 1): after the first run
	// and before the second run
	if paragraph[0].Run.Text != "Line 1" || paragraph[2].Run.Text != "Line 2" {
		t.Fatalf("the text of the runs is not preserved: %+v", paragraph)
	}
}
