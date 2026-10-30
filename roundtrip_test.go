// Copyright 2026 kenny-not-dead. All rights reserved.
// Use of this source code is governed by a BSD-style license
// that can be found in the LICENSE file.
//
// Package gopptx provides functionality to create and manipulate PowerPoint
// (.pptx) files in Go, using the Office Open XML (ECMA-376) format.

package gopptx

import (
	"archive/zip"
	"bytes"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

const testFixture = "test/presentation.pptx"

func sampleSourcePath() string {
	return filepath.Join("example", "Новая презентация.pptx")
}

// readZipPart provides a function to read a part of the zip archive.
func readZipPart(t *testing.T, path, name string) []byte {
	t.Helper()
	f, err := os.Open(path)
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()
	info, err := f.Stat()
	if err != nil {
		t.Fatal(err)
	}
	zr, err := zip.NewReader(f, info.Size())
	if err != nil {
		t.Fatal(err)
	}
	for _, v := range zr.File {
		if v.Name != name {
			continue
		}
		rc, err := v.Open()
		if err != nil {
			t.Fatal(err)
		}
		defer rc.Close()
		var buf bytes.Buffer
		if _, err := buf.ReadFrom(rc); err != nil {
			t.Fatal(err)
		}
		return buf.Bytes()
	}
	t.Fatalf("part %s not found in %s", name, path)
	return nil
}

// hasZipPart reports whether the part exists in the zip archive.
func hasZipPart(t *testing.T, path, name string) bool {
	t.Helper()
	f, err := os.Open(path)
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()
	info, err := f.Stat()
	if err != nil {
		t.Fatal(err)
	}
	zr, err := zip.NewReader(f, info.Size())
	if err != nil {
		t.Fatal(err)
	}
	for _, v := range zr.File {
		if v.Name == name {
			return true
		}
	}
	return false
}

// assertPartsEqual checks that every part of the src zip archive is present
// in the dst zip archive with the identical content.
func assertPartsEqual(t *testing.T, src, dst string) error {
	t.Helper()
	s, err := os.Open(src)
	if err != nil {
		return err
	}
	defer s.Close()
	sInfo, _ := s.Stat()
	srcReader, err := zip.NewReader(s, sInfo.Size())
	if err != nil {
		return err
	}
	d, err := os.Open(dst)
	if err != nil {
		return err
	}
	defer d.Close()
	dInfo, _ := d.Stat()
	dstReader, err := zip.NewReader(d, dInfo.Size())
	if err != nil {
		return err
	}

	parts := make(map[string]*zip.File)
	for _, v := range dstReader.File {
		parts[v.Name] = v
	}
	var errs []error
	for _, v := range srcReader.File {
		vv, ok := parts[v.Name]
		if !ok {
			errs = append(errs, fmt.Errorf("part %s is missing in the saved file", v.Name))
			continue
		}
		a, err := v.Open()
		if err != nil {
			return err
		}
		b, err := vv.Open()
		if err != nil {
			return err
		}
		var da, db bytes.Buffer
		if _, err := da.ReadFrom(a); err != nil {
			return err
		}
		if _, err := db.ReadFrom(b); err != nil {
			return err
		}
		a.Close()
		b.Close()
		if !bytes.Equal(da.Bytes(), db.Bytes()) {
			errs = append(errs, fmt.Errorf("part %s differs after the round trip (%d vs %d bytes)", v.Name, len(da.Bytes()), len(db.Bytes())))
		}
	}
	return errors.Join(errs...)
}

// TestPresentationRoundTrip checks that the test presentation keeps all of
// its content, including the images, after being opened, read and saved
// again.
func TestPresentationRoundTrip(t *testing.T) {
	f, err := OpenFile(testFixture)
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()

	// read all pictures of all slides to trigger the slide parsing
	total := 0
	for _, slideID := range f.GetSlideList() {
		pics, err := f.GetPictures(slideID)
		if err != nil {
			t.Fatal(err)
		}
		for _, pic := range pics {
			data, err := f.GetImage(slideID, pic.BlipFill.Blip.Embed)
			if err != nil {
				t.Fatal(err)
			}
			if len(data) == 0 {
				t.Fatalf("slide %d: empty image for %s", slideID, pic.BlipFill.Blip.Embed)
			}
			total++
		}
	}
	if total != 3 {
		t.Fatalf("expected 3 pictures in the test presentation, got %d", total)
	}

	out := filepath.Join(t.TempDir(), "roundtrip.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	if err := assertPartsEqual(t, testFixture, out); err != nil {
		t.Error(err)
	}
}

// TestAddDeleteImage checks that adding and then deleting an image keeps the
// presentation content unchanged.
func TestAddDeleteImage(t *testing.T) {
	f, err := OpenFile(testFixture)
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()

	// the empty slide 257
	pics, err := f.GetPictures(257)
	if err != nil {
		t.Fatal(err)
	}
	if len(pics) != 0 {
		t.Fatalf("expected no pictures on the slide 257, got %d", len(pics))
	}

	id, err := f.AddImage(257, filepath.Join("test", "images", "2.png"), nil)
	if err != nil {
		t.Fatal(err)
	}
	if err := f.DeleteImage(257, id); err != nil {
		t.Fatal(err)
	}

	out := filepath.Join(t.TempDir(), "add_delete.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	if err := assertPartsEqual(t, testFixture, out); err != nil {
		t.Fatal(err)
	}
}

// TestSamplePresentationRoundTrip checks that the sample presentation from
// the example folder keeps all of its content, including the images, after
// being opened, read and saved again.
func TestSamplePresentationRoundTrip(t *testing.T) {
	if _, err := os.Stat(sampleSourcePath()); err != nil {
		t.Skipf("sample presentation %s not found", sampleSourcePath())
	}

	f, err := OpenFile(sampleSourcePath())
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()

	// read all pictures of all slides to trigger the slide parsing
	total := 0
	for _, slideID := range f.GetSlideList() {
		pics, err := f.GetPictures(slideID)
		if err != nil {
			t.Fatal(err)
		}
		for _, pic := range pics {
			data, err := f.GetImage(slideID, pic.BlipFill.Blip.Embed)
			if err != nil {
				t.Fatal(err)
			}
			if len(data) == 0 {
				t.Fatalf("slide %d: empty image for %s", slideID, pic.BlipFill.Blip.Embed)
			}
			total++
		}
	}
	if total != 3 {
		t.Fatalf("expected 3 pictures in the sample presentation, got %d", total)
	}

	out := filepath.Join(t.TempDir(), "roundtrip.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	if err := assertPartsEqual(t, sampleSourcePath(), out); err != nil {
		t.Error(err)
	}
}

// TestSampleDeleteImage checks that deleting the pictures of the sample
// presentation (including the two pictures on the same slide which share one
// image relationship) removes the shapes and the media part, while the
// content types part keeps the origin content.
func TestSampleDeleteImage(t *testing.T) {
	if _, err := os.Stat(sampleSourcePath()); err != nil {
		t.Skipf("sample presentation %s not found", sampleSourcePath())
	}

	f, err := OpenFile(sampleSourcePath())
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()

	pics256, err := f.GetPictures(256)
	if err != nil {
		t.Fatal(err)
	}
	pics259, err := f.GetPictures(259)
	if err != nil {
		t.Fatal(err)
	}
	if len(pics256) != 1 || len(pics259) != 2 {
		t.Fatalf("unexpected sample layout: %d and %d pictures", len(pics256), len(pics259))
	}

	// the two pictures of the slide 259 share one image relationship (rId3),
	// deleting the first one must keep the relationship
	if err := f.DeleteImage(256, pics256[0].NonVisualPictureProperties.CommonNonVisualProperties.ID); err != nil {
		t.Fatal(err)
	}
	if err := f.DeleteImage(259, pics259[0].NonVisualPictureProperties.CommonNonVisualProperties.ID); err != nil {
		t.Fatal(err)
	}
	out1 := filepath.Join(t.TempDir(), "step1.pptx")
	if err := f.SaveAs(out1); err != nil {
		t.Fatal(err)
	}
	if !hasZipPart(t, out1, "ppt/media/image2.jpg") {
		t.Fatal("the shared media part was removed too early")
	}
	if n := strings.Count(string(readZipPart(t, out1, "ppt/slides/slide4.xml")), "<p:pic>"); n != 1 {
		t.Fatalf("expected 1 picture left on the slide 259, got %d", n)
	}
	if rels := string(readZipPart(t, out1, "ppt/slides/_rels/slide4.xml.rels")); !strings.Contains(rels, "relationships/image") {
		t.Fatalf("the shared image relationship of the slide 259 was removed:\n%s", rels)
	}

	// delete the last picture, the media part must be removed now
	if err := f.DeleteImage(259, pics259[1].NonVisualPictureProperties.CommonNonVisualProperties.ID); err != nil {
		t.Fatal(err)
	}
	out2 := filepath.Join(t.TempDir(), "step2.pptx")
	if err := f.SaveAs(out2); err != nil {
		t.Fatal(err)
	}
	if hasZipPart(t, out2, "ppt/media/image2.jpg") {
		t.Fatal("the media part was not removed after the last reference")
	}
	for _, rels := range []string{"ppt/slides/_rels/slide1.xml.rels", "ppt/slides/_rels/slide4.xml.rels"} {
		if s := string(readZipPart(t, out2, rels)); strings.Contains(s, "relationships/image") {
			t.Fatalf("the image relationship is not removed in %s:\n%s", rels, s)
		}
	}
	// the content types part keeps the origin content
	if !bytes.Equal(readZipPart(t, out2, "[Content_Types].xml"), readZipPart(t, sampleSourcePath(), "[Content_Types].xml")) {
		t.Fatal("the content types part differs from the origin one")
	}
}

// TestExtensionsRoundTrip checks that the standard elements the library used
// to drop on save are preserved: the extension list of the common slide data
// (p:cSld/p:extLst with the p14:creationId extension), the picture locks
// (a:picLocks), the crop rectangle of the picture (a:srcRect) and the
// dirty/err attributes of the run properties.
func TestExtensionsRoundTrip(t *testing.T) {
	fixture := filepath.Join("test", "extensions.pptx")
	f, err := OpenFile(fixture)
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()

	slideID := f.GetSlideList()[0]

	// the crop rectangle and the locks of the picture are read back
	pics, err := f.GetPictures(slideID)
	if err != nil {
		t.Fatal(err)
	}
	if len(pics) != 1 {
		t.Fatalf("expected 1 picture in the fixture, got %d", len(pics))
	}
	srcRect := pics[0].BlipFill.SrcRect
	if srcRect == nil || srcRect.Left == nil || *srcRect.Left != 10 ||
		srcRect.Top == nil || *srcRect.Top != 27 ||
		srcRect.Right == nil || *srcRect.Right != 30 ||
		srcRect.Bottom == nil || *srcRect.Bottom != 5 {
		t.Fatalf("the crop rectangle of the picture is not read back: %+v", srcRect)
	}
	locks := pics[0].NonVisualPictureProperties.CommonNonVisualPictureProperties.PicLocks
	if locks == nil || locks.NoChangeAspect == nil || !*locks.NoChangeAspect {
		t.Fatal("the noChangeAspect lock of the picture is not read back")
	}

	// the dirty/err attributes of the run properties are read back
	shapes, err := f.GetShapes(slideID)
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
			rp := body.Paragraph[j].EndParagraphRunProperties
			if rp == nil || rp.Dirty == nil || rp.Err == nil {
				continue
			}
			found = true
			if *rp.Dirty {
				t.Fatal("the dirty attribute of the run properties is not read back")
			}
			if !*rp.Err {
				t.Fatal("the err attribute of the run properties is not read back")
			}
		}
	}
	if !found {
		t.Fatal("the dirty/err attributes of the run properties are not read back")
	}

	out := filepath.Join(t.TempDir(), "extensions.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	if err := assertPartsEqual(t, fixture, out); err != nil {
		t.Error(err)
	}
}
