// Copyright 2026 kenny-not-dead. All rights reserved.
// Use of this source code is governed by a BSD-style license
// that can be found in the LICENSE file.
//
// Package gopptx provides functionality to create and manipulate PowerPoint
// (.pptx) files in Go, using the Office Open XML (ECMA-376) format.

package gopptx

import (
	"bytes"
	"encoding/xml"
	"errors"
	"path/filepath"
	"regexp"
	"strings"
	"testing"
)

// slideMasterPath provides the part path of the slide master of the new
// file.
const slideMasterTestPath = "ppt/slideMasters/slideMaster1.xml"

// layoutTestPath provides the part path of the slide layout of the new file.
const layoutTestPath = "ppt/slideLayouts/slideLayout1.xml"

// sldNumShapeBlock returns the block of the slide number placeholder shape
// (p:ph type="sldNum") of the slide layout XML: the shape is not nested in
// the shape tree, so the string splitting by the "<p:sp>" marker is safe.
func sldNumShapeBlock(layout string) (string, bool) {
	for _, part := range strings.Split(layout, "<p:sp>")[1:] {
		if strings.Contains(part, `type="sldNum"`) {
			return "<p:sp>" + part, true
		}
	}
	return "", false
}

// removeSldNumShape removes the shape of the slide number placeholder
// (p:ph type="sldNum") from the slide or the slide layout XML (the test
// fixture).
func removeSldNumShape(s string) string {
	for {
		i := strings.Index(s, `type="sldNum"`)
		if i == -1 {
			return s
		}
		start := strings.LastIndex(s[:i], "<p:sp>")
		end := strings.Index(s[i:], "</p:sp>")
		if start == -1 || end == -1 {
			return s
		}
		s = s[:start] + s[i+end+len("</p:sp>"):]
	}
}

// sldNumFieldIDRe matches the id attribute of the slide number field (the
// a:fld type="slidenum" element) of the slide or the slide layout XML.
var sldNumFieldIDRe = regexp.MustCompile(`<a:fld id="([^"]*)" type="slidenum">`)

// assertSlideSldNumField checks that the slide XML has the slide number
// field (a:fld type="slidenum") of the same id as the field of the slide
// number placeholder of the layout XML (the same as the files of Google
// Slides).
func assertSlideSldNumField(t *testing.T, slide, layout string) {
	t.Helper()
	if !strings.Contains(slide, `type="slidenum"`) {
		t.Fatalf("the slide has no the slide number field:\n%s", slide)
	}
	if m := sldNumFieldIDRe.FindStringSubmatch(layout); m != nil {
		if !strings.Contains(slide, `id="`+m[1]+`" type="slidenum"`) {
			t.Fatalf("the id of the field of the slide is not the id of the field of the layout:\n%s", slide)
		}
	}
}

// TestSetSlideNumbers checks that SetSlideNumbers sets the sldNum flag of
// the p:hf element of the slide master and keeps the slide number
// placeholder of the layout that already has the placeholder with the field
// (the layout is not modified).
func TestSetSlideNumbers(t *testing.T) {
	f := NewFile()
	defer f.Close()

	if err := f.SetSlideNumbers(true); err != nil {
		t.Fatal(err)
	}
	if !f.GetSlideNumbers() {
		t.Fatal("GetSlideNumbers = false, want true")
	}

	out := filepath.Join(t.TempDir(), "slidenumbers.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	master := string(readZipPart(t, out, slideMasterTestPath))
	if !strings.Contains(master, `sldNum="1"`) {
		t.Fatalf("the slide master has no the sldNum flag: %s", master)
	}
	// the layout of the new file already has the slide number placeholder
	// with the field: the layout is not modified on save (the part of the
	// new file is stored with the xml header, the same as on the save of the
	// unmodified file)
	origLayout := string(readZipPart(t, out, layoutTestPath))
	if origLayout != xml.Header+templateSlideLayout {
		t.Fatalf("the slide layout is modified without a reason:\n%s", origLayout)
	}
	// the slide of the new file already has the slide number placeholder
	// (the template slide): the placeholder of the slide is in the saved
	// slide
	slide := string(readZipPart(t, out, "ppt/slides/slide1.xml"))
	if !strings.Contains(slide, `type="sldNum"`) {
		t.Fatalf("the slide has no the slide number placeholder:\n%s", slide)
	}
}

// TestSetSlideNumbersSlidePlaceholderCreation checks that SetSlideNumbers
// creates the slide number placeholder (the empty placeholder, the position
// and the format of the number are inherited from the placeholder of the
// layout of the slide) on the slide that has no such placeholder (the
// template slide of the new file has no such placeholder).
func TestSetSlideNumbersSlidePlaceholderCreation(t *testing.T) {
	f := NewFile()
	defer f.Close()

	slidePath := "ppt/slides/slide1.xml"
	// the template slide is the minimal new file baseline: it has no the
	// slide number placeholder
	if slideSldNumPhRe.Match(f.readBytes(slidePath)) {
		t.Fatal("the test fixture is broken: the template slide has the slide number placeholder")
	}

	if err := f.SetSlideNumbers(true); err != nil {
		t.Fatal(err)
	}

	out := filepath.Join(t.TempDir(), "slidenumbers_slide.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	saved := string(readZipPart(t, out, slidePath))
	// the placeholder is created with the index of the placeholder of the
	// layout of the slide (the standard placeholder of the template layout)
	if !strings.Contains(saved, `type="sldNum"`) || !strings.Contains(saved, `idx="3"`) {
		t.Fatalf("the slide number placeholder is not created in the slide:\n%s", saved)
	}
	// the field of the placeholder of the slide is the field of the
	// placeholder of the layout of the slide (the same id)
	assertSlideSldNumField(t, saved, string(readZipPart(t, out, layoutTestPath)))
	// the placeholder is read back after the round trip
	f2, err := OpenFile(out)
	if err != nil {
		t.Fatal(err)
	}
	defer f2.Close()
	saved2 := string(f2.readBytes(slidePath))
	if !strings.Contains(saved2, `type="sldNum"`) {
		t.Fatalf("the slide number placeholder is lost after the round trip:\n%s", saved2)
	}
}

// TestNewSlideSlideNumberPlaceholder checks that the new slide of NewSlide
// gets the slide number placeholder when the slide numbers are explicitly
// enabled (the sldNum flag of the slide master is set to 1), and has no such
// placeholder when the numbers are not enabled.
func TestNewSlideSlideNumberPlaceholder(t *testing.T) {
	// the numbers are enabled: the new slide gets the placeholder
	f := NewFile()
	defer f.Close()
	if err := f.SetSlideNumbers(true); err != nil {
		t.Fatal(err)
	}
	slideID, err := f.NewSlide()
	if err != nil {
		t.Fatal(err)
	}
	path, ok := f.getSlideXMLPath(slideID)
	if !ok {
		t.Fatal("the new slide is not in the slide map")
	}
	out := filepath.Join(t.TempDir(), "newslide_num.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	saved := string(readZipPart(t, out, path))
	if !strings.Contains(saved, `type="sldNum"`) {
		t.Fatalf("the new slide (the numbers are enabled) has no the slide number placeholder:\n%s", saved)
	}
	// the field of the placeholder of the new slide is the field of the
	// placeholder of the layout of the slide (the same id)
	assertSlideSldNumField(t, saved, string(readZipPart(t, out, layoutTestPath)))

	// the numbers are not enabled: the new slide has no the placeholder
	f2 := NewFile()
	defer f2.Close()
	slideID2, err := f2.NewSlide()
	if err != nil {
		t.Fatal(err)
	}
	path2, ok := f2.getSlideXMLPath(slideID2)
	if !ok {
		t.Fatal("the new slide is not in the slide map")
	}
	out2 := filepath.Join(t.TempDir(), "newslide_nonum.pptx")
	if err := f2.SaveAs(out2); err != nil {
		t.Fatal(err)
	}
	saved2 := string(readZipPart(t, out2, path2))
	if slideSldNumPhRe.Match([]byte(saved2)) {
		t.Fatalf("the new slide (the numbers are not enabled) has the slide number placeholder:\n%s", saved2)
	}
}

// TestSetSlideNumbersNewSlidePlaceholder checks that SetSlideNumbers adds
// the slide number placeholder to the slide created by NewSlide before the
// call (the part of the new slide is in the in-memory store only, it is not
// in the file list of the presentation until the save).
func TestSetSlideNumbersNewSlidePlaceholder(t *testing.T) {
	f := NewFile()
	defer f.Close()

	slideID, err := f.NewSlide()
	if err != nil {
		t.Fatal(err)
	}
	path, ok := f.getSlideXMLPath(slideID)
	if !ok {
		t.Fatal("the new slide is not in the slide map")
	}
	if err := f.SetSlideNumbers(true); err != nil {
		t.Fatal(err)
	}
	out := filepath.Join(t.TempDir(), "slidenumbers_newslide.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	saved := string(readZipPart(t, out, path))
	if !strings.Contains(saved, `type="sldNum"`) {
		t.Fatalf("the new slide has no the slide number placeholder:\n%s", saved)
	}
	// the field of the placeholder of the new slide is the field of the
	// placeholder of the layout of the slide (the same id)
	assertSlideSldNumField(t, saved, string(readZipPart(t, out, layoutTestPath)))
}

// TestSetSlideNumbersOptions checks that the options of SetSlideNumbers are
// applied to the slide number placeholder of the layout: the position of
// the placeholder and the properties of the number text (the size, the
// color and the typeface of the field).
func TestSetSlideNumbersOptions(t *testing.T) {
	f := NewFile()
	defer f.Close()

	size, bold := 1800, 1
	opts := SlideNumberOptions{
		Xfrm: &DecodeXfrm{
			Offset:  &Offset{X: 8709025, Y: 4756150},
			Extents: &Extents{CX: 914400, CY: 457200},
		},
		Run: &DecodeRunProperties{
			Size:  &size,
			Bold:  &bold,
			Latin: &Font{Typeface: "Arial"},
			SolidFill: &DecodeSolidFill{
				SolidRGBColor: &SolidRGBColor{Val: "FF0000"},
			},
		},
	}
	if err := f.SetSlideNumbers(true, opts); err != nil {
		t.Fatal(err)
	}

	out := filepath.Join(t.TempDir(), "slidenumbers_opts.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	layout := string(readZipPart(t, out, layoutTestPath))
	shape, ok := sldNumShapeBlock(layout)
	if !ok {
		t.Fatalf("the slide layout has no the slide number placeholder:\n%s", layout)
	}
	for _, want := range []string{
		`<a:off x="8709025" y="4756150"/>`,
		`<a:ext cx="914400" cy="457200"/>`,
		`sz="1800"`,
		`b="1"`,
		`typeface="Arial"`,
		`<a:srgbClr val="FF0000"/>`,
	} {
		if !strings.Contains(shape, want) {
			t.Errorf("the slide number placeholder has no %s:\n%s", want, shape)
		}
	}
}

// TestSetSlideNumbersPlaceholderCreation checks that SetSlideNumbers creates
// the slide number placeholder (with the slide number field) in the layout
// that has no such placeholder, keeping the other shapes of the layout.
func TestSetSlideNumbersPlaceholderCreation(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// the slide layout of the new file without the slide number placeholder
	// (the same as the slide layouts of the real presentation that has no
	// the slide number placeholder)
	layout := removeSldNumShape(templateSlideLayout)
	if layout == templateSlideLayout {
		t.Fatal("the test fixture is broken: the placeholder is not removed")
	}
	f.Pkg.Store(layoutTestPath, []byte(layout))

	if err := f.SetSlideNumbers(true); err != nil {
		t.Fatal(err)
	}

	out := filepath.Join(t.TempDir(), "slidenumbers_new.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	saved := string(readZipPart(t, out, layoutTestPath))
	shape, ok := sldNumShapeBlock(saved)
	if !ok {
		t.Fatalf("the slide number placeholder is not created in the layout:\n%s", saved)
	}
	for _, want := range []string{
		`<p:ph idx="12" type="sldNum"/>`,
		`type="slidenum"`,
		`Slide Number Placeholder 12`,
	} {
		if !strings.Contains(shape, want) {
			t.Errorf("the created slide number placeholder has no %s:\n%s", want, shape)
		}
	}
	// the other placeholders of the layout are kept
	for _, want := range []string{`<p:ph type="title"/>`, `<p:ph type="subTitle"/>`} {
		if !strings.Contains(saved, want) {
			t.Errorf("the slide layout has lost the placeholder %s:\n%s", want, saved)
		}
	}
	// the placeholder with the field is read back
	f2, err := OpenFile(out)
	if err != nil {
		t.Fatal(err)
	}
	defer f2.Close()
	if !f2.GetSlideNumbers() {
		t.Fatal("GetSlideNumbers after the round trip = false, want true")
	}
}

// TestSetSlideNumbersDisable checks that SetSlideNumbers with enable=false
// sets the sldNum flag of the slide master to 0 and keeps the placeholder
// in the layout.
func TestSetSlideNumbersDisable(t *testing.T) {
	f := NewFile()
	defer f.Close()

	if err := f.SetSlideNumbers(true); err != nil {
		t.Fatal(err)
	}
	if err := f.SetSlideNumbers(false); err != nil {
		t.Fatal(err)
	}
	if f.GetSlideNumbers() {
		t.Fatal("GetSlideNumbers = true after the disable, want false")
	}

	out := filepath.Join(t.TempDir(), "slidenumbers_off.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	master := string(readZipPart(t, out, slideMasterTestPath))
	if !strings.Contains(master, `sldNum="0"`) {
		t.Fatalf("the slide master has no the sldNum flag of the value 0: %s", master)
	}
	layout := string(readZipPart(t, out, layoutTestPath))
	if !strings.Contains(layout, `type="sldNum"`) {
		t.Fatalf("the slide layout has lost the slide number placeholder: %s", layout)
	}
}

// TestSetSlideNumbersRoundTrip checks that the parts modified by
// SetSlideNumbers (the slide master and the slide layout) are stable on the
// next save: the file saved after the round trip is identical to the file
// saved right after the modification.
func TestSetSlideNumbersRoundTrip(t *testing.T) {
	f := NewFile()
	defer f.Close()

	size := 1800
	if err := f.SetSlideNumbers(true, SlideNumberOptions{
		Xfrm: &DecodeXfrm{
			Offset:  &Offset{X: 8709025, Y: 4756150},
			Extents: &Extents{CX: 914400, CY: 457200},
		},
		Run: &DecodeRunProperties{Size: &size},
	}); err != nil {
		t.Fatal(err)
	}

	first := filepath.Join(t.TempDir(), "slidenumbers_1.pptx")
	if err := f.SaveAs(first); err != nil {
		t.Fatal(err)
	}

	f2, err := OpenFile(first)
	if err != nil {
		t.Fatal(err)
	}
	defer f2.Close()

	second := filepath.Join(t.TempDir(), "slidenumbers_2.pptx")
	if err := f2.SaveAs(second); err != nil {
		t.Fatal(err)
	}
	for _, part := range []string{slideMasterTestPath, layoutTestPath, "ppt/slides/slide1.xml"} {
		a := readZipPart(t, first, part)
		b := readZipPart(t, second, part)
		if !bytes.Equal(a, b) {
			t.Errorf("%s differs after the round trip:\n--- first\n%s\n--- second\n%s", part, a, b)
		}
	}
}

// TestSetSlideNumbersUnsupportedLayout checks that SetSlideNumbers returns
// the ErrSlideLayoutUnsupported error when the slide layout contains an
// element that is not supported (the picture of the shape tree).
func TestSetSlideNumbersUnsupportedLayout(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// the slide layout with the picture in the shape tree
	layout := strings.Replace(templateSlideLayout, "</p:spTree>",
		`<p:pic><p:nvPicPr><p:cNvPr id="100" name="pic"/></p:nvPicPr></p:pic></p:spTree>`, 1)
	f.Pkg.Store(layoutTestPath, []byte(layout))

	err := f.SetSlideNumbers(true)
	if !errors.As(err, new(ErrSlideLayoutUnsupported)) {
		t.Fatalf("expected ErrSlideLayoutUnsupported, got %v", err)
	}
}
