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
	"os"
	"path/filepath"
	"strings"
	"testing"
)

// testNotesText returns the test notes text body.
func testNotesText() DecodeTextBody {
	bold := 1
	return DecodeTextBody{
		BodyProperties: &DecodeBodyProperties{},
		Paragraph: []DecodeParagraph{
			{
				Runs: []DecodeRuns{{Text: "Notes for the first slide."}},
			},
			{
				Runs: []DecodeRuns{{
					RunProperties: &DecodeRunProperties{Bold: &bold},
					Text:          "The second paragraph.",
				}},
			},
		},
	}
}

// textBodyString returns the text of the text body: the text of the runs of
// the paragraphs separated by a line break.
func textBodyString(tb DecodeTextBody) string {
	var sb strings.Builder
	for _, p := range tb.Paragraph {
		for _, r := range p.Runs {
			sb.WriteString(r.Text)
		}
		sb.WriteString("\n")
	}

	return sb.String()
}

// TestNewFileSetNotes checks that SetNotes creates the notes slide and the
// notes master in the package of the new file and that the notes are
// preserved after the round trip.
func TestNewFileSetNotes(t *testing.T) {
	f := NewFile()
	defer f.Close()
	slideID := defaultXMLSlideID

	if _, err := f.GetNotes(slideID); !errors.As(err, new(ErrNotesSlideNotExist)) {
		t.Fatalf("expected ErrNotesSlideNotExist, got %v", err)
	}

	if err := f.SetNotes(slideID, testNotesText()); err != nil {
		t.Fatal(err)
	}
	got, err := f.GetNotes(slideID)
	if err != nil {
		t.Fatal(err)
	}
	if want := textBodyString(testNotesText()); textBodyString(got) != want {
		t.Fatalf("the notes text is %q, want %q", textBodyString(got), want)
	}

	out := filepath.Join(t.TempDir(), "notes.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	for _, name := range []string{
		"ppt/notesSlides/notesSlide1.xml",
		"ppt/notesSlides/_rels/notesSlide1.xml.rels",
		"ppt/notesMasters/notesMaster1.xml",
		"ppt/notesMasters/_rels/notesMaster1.xml.rels",
		"ppt/theme/theme2.xml",
	} {
		if !hasZipPart(t, out, name) {
			t.Errorf("the part %s is missing in the saved file", name)
		}
	}
	pres := readZipPart(t, out, "ppt/presentation.xml")
	if !strings.Contains(string(pres), "notesMasterIdLst") {
		t.Error("the presentation.xml has no notesMasterIdLst element")
	}

	f2, err := OpenFile(out)
	if err != nil {
		t.Fatal(err)
	}
	defer f2.Close()
	got2, err := f2.GetNotes(slideID)
	if err != nil {
		t.Fatal(err)
	}
	if want := textBodyString(testNotesText()); textBodyString(got2) != want {
		t.Fatalf("the notes text after the round trip is %q, want %q", textBodyString(got2), want)
	}
}

// TestNewNotes checks that NewNotes creates an empty notes slide,
// that it returns an error when the notes slide already exists and for the
// non-existing slide.
func TestNewNotes(t *testing.T) {
	f := NewFile()
	defer f.Close()
	slideID := defaultXMLSlideID

	if err := f.NewNotes(999); !errors.As(err, new(ErrSlideNotExist)) {
		t.Fatalf("expected ErrSlideNotExist, got %v", err)
	}
	if err := f.NewNotes(slideID); err != nil {
		t.Fatal(err)
	}
	if err := f.NewNotes(slideID); !errors.As(err, new(ErrNotesSlideExists)) {
		t.Fatalf("expected ErrNotesSlideExists, got %v", err)
	}

	got, err := f.GetNotes(slideID)
	if err != nil {
		t.Fatal(err)
	}
	if textBodyString(got) != "\n" {
		t.Fatalf("the new notes slide must be empty, got %q", textBodyString(got))
	}
}

// TestDeleteNotes checks that DeleteNotes removes the notes slide
// part, its rels part and the notesSlide relationship of the slide.
func TestDeleteNotes(t *testing.T) {
	f := NewFile()
	defer f.Close()
	slideID := defaultXMLSlideID

	if err := f.DeleteNotes(999); !errors.As(err, new(ErrNotesSlideNotExist)) {
		t.Fatalf("expected ErrNotesSlideNotExist, got %v", err)
	}
	if err := f.NewNotes(slideID); err != nil {
		t.Fatal(err)
	}
	if err := f.DeleteNotes(slideID); err != nil {
		t.Fatal(err)
	}
	if _, err := f.GetNotes(slideID); !errors.As(err, new(ErrNotesSlideNotExist)) {
		t.Fatalf("expected ErrNotesSlideNotExist, got %v", err)
	}
	if err := f.DeleteNotes(slideID); !errors.As(err, new(ErrNotesSlideNotExist)) {
		t.Fatalf("expected ErrNotesSlideNotExist, got %v", err)
	}

	out := filepath.Join(t.TempDir(), "notes_deleted.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	if hasZipPart(t, out, "ppt/notesSlides/notesSlide1.xml") {
		t.Error("the notes slide part is still in the saved file")
	}
	if hasZipPart(t, out, "ppt/notesSlides/_rels/notesSlide1.xml.rels") {
		t.Error("the notes slide rels part is still in the saved file")
	}
	slideRels := string(readZipPart(t, out, "ppt/slides/_rels/slide1.xml.rels"))
	if strings.Contains(slideRels, "notesSlide") {
		t.Errorf("the slide rels part still references the notes slide: %s", slideRels)
	}
}

// TestSampleNotesRoundTrip checks that the notes slide of the sample
// presentation (the Google Slides file) is preserved byte-by-byte after it
// is loaded by GetNotes and the file is saved.
func TestSampleNotesRoundTrip(t *testing.T) {
	src := sampleSourcePath()
	f, err := OpenFile(src)
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()

	// find a slide with the notes
	slideID := 0
	for _, id := range f.GetSlideList() {
		if _, err := f.GetNotes(id); err == nil {
			slideID = id
			break
		}
	}
	if slideID == 0 {
		t.Fatal("the sample presentation has no notes")
	}
	got, err := f.GetNotes(slideID)
	if err != nil {
		t.Fatal(err)
	}

	out := filepath.Join(t.TempDir(), "sample_notes.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}

	// the loaded notes slide must be serialized byte-by-byte identical
	srcNotes := readZipPart(t, src, "ppt/notesSlides/notesSlide1.xml")
	dstNotes := readZipPart(t, out, "ppt/notesSlides/notesSlide1.xml")
	if !bytes.Equal(srcNotes, dstNotes) {
		t.Errorf("the notes slide differs after the round trip:\n--- src\n%s\n--- dst\n%s", srcNotes, dstNotes)
	}

	f2, err := OpenFile(out)
	if err != nil {
		t.Fatal(err)
	}
	defer f2.Close()
	got2, err := f2.GetNotes(slideID)
	if err != nil {
		t.Fatal(err)
	}
	if textBodyString(got) != textBodyString(got2) {
		t.Fatalf("the notes text after the round trip is %q, want %q", textBodyString(got2), textBodyString(got))
	}
}

// TestDeleteSlideRemovesNotes checks that DeleteSlide removes the notes slide
// part of the deleted slide from the package.
func TestDeleteSlideRemovesNotes(t *testing.T) {
	src := sampleSourcePath()
	f, err := OpenFile(src)
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()

	ids := f.GetSlideList()
	if len(ids) < 2 {
		t.Skip("the sample presentation has no extra slides")
	}
	victim := ids[0]
	if _, err := f.GetNotes(victim); err != nil {
		t.Skip("the first slide of the sample has no notes")
	}
	notesPath, ok := f.getNotesSlidePath(victim)
	if !ok {
		t.Skip("the first slide of the sample has no notes")
	}

	if err := f.DeleteSlide(victim); err != nil {
		t.Fatal(err)
	}
	out := filepath.Join(t.TempDir(), "slide_with_notes_deleted.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	if hasZipPart(t, out, notesPath) {
		t.Error("the notes slide part of the deleted slide is still in the package")
	}
	notesRelsPath := "ppt/notesSlides/_rels/" + filepath.Base(notesPath) + ".rels"
	if hasZipPart(t, out, notesRelsPath) {
		t.Error("the notes slide rels part of the deleted slide is still in the package")
	}
}

// TestSetShapeHyperLink checks that SetShapeHyperLink sets the hyperlink to
// the runs of the shape, adds the hyperlink relationship to the slide rels
// part and that the hyperlink is preserved after the round trip.
func TestSetShapeHyperLink(t *testing.T) {
	f := NewFile()
	defer f.Close()
	slideID := defaultXMLSlideID
	shapeID := defaultXMLShapeID
	link := "https://github.com/kenny-not-dead/gopptx"
	tooltip := "gopptx on GitHub"

	if err := f.SetShapeHyperLink(slideID, 999, link, LinkTypeExternal); !errors.As(err, new(ErrShapeNotExist)) {
		t.Fatalf("expected ErrShapeNotExist, got %v", err)
	}
	if err := f.SetShapeHyperLink(slideID, shapeID, link, LinkType("Invalid")); !errors.Is(err, ErrInvalidLinkType) {
		t.Fatalf("expected ErrInvalidLinkType, got %v", err)
	}

	if err := f.SetShapeHyperLink(slideID, shapeID, link, LinkTypeExternal, HyperlinkOpts{Tooltip: &tooltip}); err != nil {
		t.Fatal(err)
	}
	hasLink, got, err := f.GetShapeHyperLink(slideID, shapeID)
	if err != nil {
		t.Fatal(err)
	}
	if !hasLink || got != link {
		t.Fatalf("GetShapeHyperLink = (%v, %q), want (true, %q)", hasLink, got, link)
	}

	ids, err := f.GetShapeHyperLinks(slideID, "")
	if err != nil {
		t.Fatal(err)
	}
	if len(ids) != 1 || ids[0] != shapeID {
		t.Fatalf("GetShapeHyperLinks = %v, want [%d]", ids, shapeID)
	}
	ids, err = f.GetShapeHyperLinks(slideID, LinkTypeExternal)
	if err != nil {
		t.Fatal(err)
	}
	if len(ids) != 1 || ids[0] != shapeID {
		t.Fatalf("GetShapeHyperLinks External = %v, want [%d]", ids, shapeID)
	}
	if _, err := f.GetShapeHyperLinks(slideID, LinkType("Invalid")); !errors.Is(err, ErrInvalidLinkType) {
		t.Fatalf("expected ErrInvalidLinkType, got %v", err)
	}

	out := filepath.Join(t.TempDir(), "hyperlink.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	slideXML := string(readZipPart(t, out, "ppt/slides/slide1.xml"))
	if !strings.Contains(slideXML, "hlinkClick") {
		t.Errorf("the slide has no hlinkClick element: %s", slideXML)
	}
	if !strings.Contains(slideXML, `tooltip="gopptx on GitHub"`) {
		t.Errorf("the hlinkClick element has no tooltip: %s", slideXML)
	}
	slideRels := string(readZipPart(t, out, "ppt/slides/_rels/slide1.xml.rels"))
	if !strings.Contains(slideRels, SourceRelationshipHyperlink) {
		t.Errorf("the slide rels have no hyperlink relationship: %s", slideRels)
	}
	if !strings.Contains(slideRels, `TargetMode="External"`) {
		t.Errorf("the hyperlink relationship has no TargetMode External: %s", slideRels)
	}

	f2, err := OpenFile(out)
	if err != nil {
		t.Fatal(err)
	}
	defer f2.Close()
	hasLink, got, err = f2.GetShapeHyperLink(slideID, shapeID)
	if err != nil {
		t.Fatal(err)
	}
	if !hasLink || got != link {
		t.Fatalf("GetShapeHyperLink after the round trip = (%v, %q), want (true, %q)", hasLink, got, link)
	}
}

// TestSetShapeHyperLinkNone checks that SetShapeHyperLink with the link type
// LinkTypeNone removes the hyperlink from the shape and the hyperlink
// relationship from the slide rels part.
func TestSetShapeHyperLinkNone(t *testing.T) {
	f := NewFile()
	defer f.Close()
	slideID := defaultXMLSlideID
	shapeID := defaultXMLShapeID
	link := "https://github.com/kenny-not-dead/gopptx"

	if err := f.SetShapeHyperLink(slideID, shapeID, link, LinkTypeExternal); err != nil {
		t.Fatal(err)
	}
	if err := f.SetShapeHyperLink(slideID, shapeID, "", LinkTypeNone); err != nil {
		t.Fatal(err)
	}
	hasLink, got, err := f.GetShapeHyperLink(slideID, shapeID)
	if err != nil {
		t.Fatal(err)
	}
	if hasLink || got != "" {
		t.Fatalf("GetShapeHyperLink = (%v, %q), want (false, \"\")", hasLink, got)
	}

	out := filepath.Join(t.TempDir(), "hyperlink_none.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	slideXML := string(readZipPart(t, out, "ppt/slides/slide1.xml"))
	if strings.Contains(slideXML, "hlinkClick") {
		t.Errorf("the slide still has the hlinkClick element: %s", slideXML)
	}
	slideRels := string(readZipPart(t, out, "ppt/slides/_rels/slide1.xml.rels"))
	if strings.Contains(slideRels, SourceRelationshipHyperlink) {
		t.Errorf("the slide rels still have the hyperlink relationship: %s", slideRels)
	}
}

// TestSetShapeHyperLinkPicture checks that SetShapeHyperLink sets the
// click-through hyperlink to the common non-visual properties of the picture.
func TestSetShapeHyperLinkPicture(t *testing.T) {
	f := NewFile()
	defer f.Close()
	slideID := defaultXMLSlideID
	link := "https://github.com/kenny-not-dead/gopptx"

	image, err := os.ReadFile("test/image1.jpg")
	if err != nil {
		t.Fatal(err)
	}
	pictureID, err := f.AddImageFromBytes(slideID, image, nil)
	if err != nil {
		t.Fatal(err)
	}

	if err := f.SetShapeHyperLink(slideID, pictureID, link, LinkTypeExternal); err != nil {
		t.Fatal(err)
	}
	hasLink, got, err := f.GetShapeHyperLink(slideID, pictureID)
	if err != nil {
		t.Fatal(err)
	}
	if !hasLink || got != link {
		t.Fatalf("GetShapeHyperLink = (%v, %q), want (true, %q)", hasLink, got, link)
	}

	out := filepath.Join(t.TempDir(), "hyperlink_picture.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	slideXML := string(readZipPart(t, out, "ppt/slides/slide1.xml"))
	if !strings.Contains(slideXML, "hlinkClick") {
		t.Errorf("the picture has no hlinkClick element: %s", slideXML)
	}

	f2, err := OpenFile(out)
	if err != nil {
		t.Fatal(err)
	}
	defer f2.Close()
	hasLink, got, err = f2.GetShapeHyperLink(slideID, pictureID)
	if err != nil {
		t.Fatal(err)
	}
	if !hasLink || got != link {
		t.Fatalf("GetShapeHyperLink after the round trip = (%v, %q), want (true, %q)", hasLink, got, link)
	}
}
