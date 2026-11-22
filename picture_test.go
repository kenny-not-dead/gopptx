// Copyright 2026 kenny-not-dead. All rights reserved.
// Use of this source code is governed by a BSD-style license
// that can be found in the LICENSE file.
//
// Package gopptx provides functionality to create and manipulate PowerPoint
// (.pptx) files in Go, using the Office Open XML (ECMA-376) format.

package gopptx

import (
	"fmt"
	"image"
	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"
	"io"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/stretchr/testify/assert"
)

// macCyrillicCharset is a non UTF-8 bytes sequence used to simulate a
// presentation part written with an unsupported charset.
var macCyrillicCharset = []byte("\xC1\xC2\xC3")

func BenchmarkAddImageFromBytes(b *testing.B) {
	f := NewFile()
	slideID := f.GetSlideList()[0]
	imgFile, err := os.ReadFile(filepath.Join("test", "images", "1.png"))
	if err != nil {
		b.Error("unable to load image for benchmark")
	}
	b.ResetTimer()
	for i := 1; i <= b.N; i++ {
		if _, err := f.AddImageFromBytes(slideID, imgFile, &PictureOptions{Name: fmt.Sprintf("Image %d", i)}); err != nil {
			b.Error(err)
		}
	}
}

func TestAddImage(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "presentation.pptx"))
	assert.NoError(t, err)

	// Test add image to a slide with the name option
	opts := []PictureOptions{
		{Name: "Image 4"},
		{Name: "Image 4", OffsetX: 457200, OffsetY: 381000, ScaleX: 0.5, ScaleY: 0.5},
		{Name: "Image 5", OffsetX: 457200, OffsetY: 381000, ScaleX: 0.88, ScaleY: 0.88},
		{Name: "Image 4", ScaleX: 1, ScaleY: 1},
		{Name: "Image 9"},
	}
	_, err = f.AddImage(257, filepath.Join("test", "images", "1.jpg"), &opts[0])
	assert.NoError(t, err)
	// Test add image to a slide with offset and scale
	_, err = f.AddImage(256, filepath.Join("test", "images", "1.jpg"), &opts[1])
	assert.NoError(t, err)
	_, err = f.AddImage(256, filepath.Join("test", "images", "1.jpg"), &opts[2])
	assert.NoError(t, err)

	// Test add images to the same slide with offsets
	_, err = f.AddImage(258, filepath.Join("test", "images", "1.jpg"), nil)
	assert.NoError(t, err)
	_, err = f.AddImage(258, filepath.Join("test", "images", "1.jpg"), &opts[3])
	assert.NoError(t, err)
	_, err = f.AddImage(258, filepath.Join("test", "images", "1.jpg"), &PictureOptions{OffsetX: 914400})
	assert.NoError(t, err)
	_, err = f.AddImage(258, filepath.Join("test", "images", "1.jpg"), &PictureOptions{OffsetX: 1828800})
	assert.NoError(t, err)

	file, err := os.ReadFile(filepath.Join("test", "images", "1.png"))
	assert.NoError(t, err)

	// Test add image to a slide with the default layout
	_, err = f.AddImage(257, filepath.Join("test", "images", "2.png"), nil)
	assert.NoError(t, err)
	_, err = f.AddImage(257, filepath.Join("test", "images", "2.png"), &PictureOptions{OffsetX: 457200, OffsetY: 381000})
	assert.NoError(t, err)
	_, err = f.AddImage(257, filepath.Join("test", "images", "2.png"), &PictureOptions{ScaleX: 0.5, ScaleY: 0.5})
	assert.NoError(t, err)
	// Test add image to a new slide
	newSlideID, err := f.NewSlide()
	assert.NoError(t, err)
	_, err = f.AddImage(newSlideID, filepath.Join("test", "images", "2.png"), nil)
	assert.NoError(t, err)

	// Test add image to a slide from bytes
	_, err = f.AddImageFromBytes(newSlideID, file, &opts[4])
	assert.NoError(t, err)
	// Test add image to a not exists slide from bytes
	_, err = f.AddImageFromBytes(999, file, nil)
	assert.EqualError(t, err, ErrSlideNotExist{999}.Error())

	for _, preset := range []struct {
		slideID int
		ext     string
	}{
		{256, "gif"}, {257, "tiff"}, {258, "bmp"}, {257, "jpg"},
	} {
		_, err = f.AddImage(preset.slideID, filepath.Join("test", "images", "1."+preset.ext), nil)
		assert.NoError(t, err)
	}

	// Test get pictures from the slide which added the image with the name option
	pics, err := f.GetPictures(257)
	assert.NoError(t, err)
	assert.Len(t, pics, 6)
	assert.Equal(t, opts[0].Name, pics[0].NonVisualPictureProperties.CommonNonVisualProperties.Name)
	// Test get pictures from the slide which added the images with offset and scale
	pics, err = f.GetPictures(256)
	assert.NoError(t, err)
	assert.Len(t, pics, 4)
	assert.Equal(t, opts[1].Name, pics[1].NonVisualPictureProperties.CommonNonVisualProperties.Name)
	assert.Equal(t, opts[1].OffsetX, pics[1].ShapeProperties.Xfrm.Offset.X)
	assert.Equal(t, opts[1].OffsetY, pics[1].ShapeProperties.Xfrm.Offset.Y)
	assert.Equal(t, 1524000, pics[1].ShapeProperties.Xfrm.Extents.CX)
	assert.Equal(t, 952500, pics[1].ShapeProperties.Xfrm.Extents.CY)
	assert.Equal(t, opts[2].Name, pics[2].NonVisualPictureProperties.CommonNonVisualProperties.Name)
	assert.Equal(t, 2682240, pics[2].ShapeProperties.Xfrm.Extents.CX)
	assert.Equal(t, 1676400, pics[2].ShapeProperties.Xfrm.Extents.CY)
	// Test get pictures from the slide which added the images with offsets
	pics, err = f.GetPictures(258)
	assert.NoError(t, err)
	assert.Len(t, pics, 7)
	assert.Equal(t, opts[3].Name, pics[3].NonVisualPictureProperties.CommonNonVisualProperties.Name)
	assert.Equal(t, 914400, pics[4].ShapeProperties.Xfrm.Offset.X)
	assert.Equal(t, 0, pics[4].ShapeProperties.Xfrm.Offset.Y)
	// Test get pictures from the new slide which added the image from bytes
	pics, err = f.GetPictures(newSlideID)
	assert.NoError(t, err)
	assert.Len(t, pics, 2)
	assert.Equal(t, opts[4].Name, pics[1].NonVisualPictureProperties.CommonNonVisualProperties.Name)

	// The duplicated images are only stored once in the media folder
	imageCount := 0
	f.Pkg.Range(func(fileName, v interface{}) bool {
		if strings.Contains(fileName.(string), "media/image") {
			imageCount++
		}
		return true
	})
	assert.Equal(t, 6, imageCount, "Duplicate image should only be stored once.")

	// Test write file to given path
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestAddImage1.pptx")))
	assert.NoError(t, f.Close())

	// Test get pictures after inserting a new image from a presentation which
	// contains existing images
	f, err = OpenFile(filepath.Join("test", "TestAddImage1.pptx"))
	assert.NoError(t, err)
	_, err = f.AddImage(256, filepath.Join("test", "images", "1.jpg"), nil)
	assert.NoError(t, err)
	pics, err = f.GetPictures(256)
	assert.NoError(t, err)
	assert.Len(t, pics, 5)

	// Test get pictures from the slides which already contain images
	pics, err = f.GetPictures(258)
	assert.NoError(t, err)
	assert.Len(t, pics, 7)
	assert.Equal(t, opts[3].Name, pics[3].NonVisualPictureProperties.CommonNonVisualProperties.Name)
	pics, err = f.GetPictures(257)
	assert.NoError(t, err)
	assert.Len(t, pics, 6)
	assert.Equal(t, opts[0].Name, pics[0].NonVisualPictureProperties.CommonNonVisualProperties.Name)
	pics, err = f.GetPictures(256)
	assert.NoError(t, err)
	assert.Len(t, pics, 5)
	assert.Equal(t, opts[1].Name, pics[1].NonVisualPictureProperties.CommonNonVisualProperties.Name)
	assert.NoError(t, f.Close())

	// Test add image with unsupported charset content types
	f = NewFile()
	f.ContentTypes = nil
	f.Pkg.Store(defaultXMLPathContentTypes, macCyrillicCharset)
	_, err = f.AddImageFromBytes(f.GetSlideList()[0], file, nil)
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
	assert.NoError(t, f.Close())
}

func TestAddImageErrors(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "presentation.pptx"))
	assert.NoError(t, err)

	// Test add image to a slide with invalid file path
	_, err = f.AddImage(256, filepath.Join("test", "not_exists_dir", "not_exists.icon"), nil)
	assert.Error(t, err)

	// Test add image to a slide with unsupported file type
	_, err = f.AddImage(256, filepath.Join("test", "presentation.pptx"), nil)
	assert.EqualError(t, err, ErrUnsupportedImageFormat.Error())

	// Test add image to a slide from bytes with invalid file data
	_, err = f.AddImageFromBytes(256, []byte{1}, nil)
	assert.EqualError(t, err, ErrUnsupportedImageFormat.Error())

	// Test add image to a slide from bytes with empty data
	_, err = f.AddImageFromBytes(256, nil, nil)
	assert.EqualError(t, err, ErrImageEmpty.Error())

	// Test add image with custom image decoder and encoder
	decode := func(r io.Reader) (image.Image, error) { return nil, nil }
	decodeConfig := func(r io.Reader) (image.Config, error) { return image.Config{Height: 128, Width: 200}, nil }
	for _, preset := range []struct {
		slideID int
		ext     string
	}{
		{256, "emf"}, {258, "wmf"},
	} {
		image.RegisterFormat(preset.ext, "", decode, decodeConfig)
		_, err = f.AddImage(preset.slideID, filepath.Join("test", "images", "1."+preset.ext), &PictureOptions{ScaleX: 0.8, ScaleY: 0.8})
		assert.NoError(t, err)
	}
	_, err = f.AddImage(257, filepath.Join("test", "images", "1.svg"), &PictureOptions{ScaleX: 1.6})
	assert.NoError(t, err)
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestAddImage2.pptx")))
	assert.NoError(t, f.Close())
}

func TestGetImage(t *testing.T) {
	f := NewFile()
	_, err := f.AddImage(256, filepath.Join("test", "images", "1.png"), nil)
	assert.NoError(t, err)
	pics, err := f.GetPictures(256)
	assert.NoError(t, err)
	assert.Len(t, pics, 1)
	data, err := f.GetImage(256, pics[0].BlipFill.Blip.Embed)
	assert.NoError(t, err)
	imgFile, err := os.ReadFile(filepath.Join("test", "images", "1.png"))
	assert.NoError(t, err)
	assert.Len(t, data, len(imgFile))
	assert.Equal(t, "Picture 9", pics[0].NonVisualPictureProperties.CommonNonVisualProperties.Name)
	assert.NoError(t, f.Close())

	f, err = OpenFile(filepath.Join("test", "presentation.pptx"))
	assert.NoError(t, err)
	// Try to get pictures from a slide that doesn't contain any images
	pics, err = f.GetPictures(257)
	assert.NoError(t, err)
	assert.Len(t, pics, 0)
	assert.NoError(t, f.Close())

	f, err = prepareTestPresentation()
	if !assert.NoError(t, err) {
		t.FailNow()
	}

	// The image added by prepareTestPresentation is readable and its data can
	// be written to a local file
	pics, err = f.GetPictures(257)
	assert.NoError(t, err)
	assert.Len(t, pics, 1)
	data, err = f.GetImage(257, pics[0].BlipFill.Blip.Embed)
	assert.NoError(t, err)
	if !assert.NotEmpty(t, data) || !assert.NoError(t, os.WriteFile(filepath.Join(t.TempDir(), "image1.jpg"), data, 0o644)) {
		t.FailNow()
	}

	// Try to get picture with not exists picture id
	_, err = f.GetPicture(257, 999)
	assert.Equal(t, ErrPictureNotExist{999}, err)
	// Try to get image with not exists relationship id
	_, err = f.GetImage(257, "rId99")
	assert.Equal(t, ErrImageNotExist{"rId99"}, err)

	// Test write file to given path
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestGetImage.pptx")))
	assert.NoError(t, f.Close())

	f, err = OpenFile(filepath.Join("test", "TestGetImage.pptx"))
	assert.NoError(t, err)
	// Test get pictures from the slide which already contains images added
	// from bytes
	pics, err = f.GetPictures(258)
	assert.NoError(t, err)
	assert.Len(t, pics, 3)
	assert.Equal(t, "Logo", pics[2].NonVisualPictureProperties.CommonNonVisualProperties.Name)
	data, err = f.GetImage(258, pics[2].BlipFill.Blip.Embed)
	assert.NoError(t, err)
	assert.Equal(t, imgFile, data)
	// Try to get image with the relationship id of another slide
	_, err = f.GetImage(257, pics[2].BlipFill.Blip.Embed)
	assert.Equal(t, ErrImageNotExist{pics[2].BlipFill.Blip.Embed}, err)
	assert.NoError(t, f.Close())

	// Test get pictures with unsupported charset slide
	f, err = OpenFile(filepath.Join("test", "TestGetImage.pptx"))
	assert.NoError(t, err)
	f.Slide.Delete("ppt/slides/slide3.xml")
	f.Pkg.Store("ppt/slides/slide3.xml", macCyrillicCharset)
	_, err = f.GetPictures(258)
	assert.EqualError(t, err, "XML syntax error on line 1: invalid UTF-8")
	assert.NoError(t, f.Close())
}

func TestAddImageFromBytes(t *testing.T) {
	f := NewFile()
	slideID := f.GetSlideList()[0]
	imgFile, err := os.ReadFile(filepath.Join("test", "images", "1.png"))
	assert.NoError(t, err, "Unable to load image for test")

	_, err = f.AddImageFromBytes(slideID, imgFile, &PictureOptions{Name: "Image 1"})
	assert.NoError(t, err)
	_, err = f.AddImageFromBytes(slideID, imgFile, &PictureOptions{Name: "Image 2"})
	assert.NoError(t, err)
	imageCount := 0
	f.Pkg.Range(func(fileName, v interface{}) bool {
		if strings.Contains(fileName.(string), "media/image") {
			imageCount++
		}
		return true
	})
	assert.Equal(t, 1, imageCount, "Duplicate image should only be stored once.")
	// Test add image from bytes on not exists slide
	_, err = f.AddImageFromBytes(999, imgFile, nil)
	assert.EqualError(t, err, ErrSlideNotExist{999}.Error())
	assert.NoError(t, f.Close())
}

func TestDeleteImage(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "presentation.pptx"))
	assert.NoError(t, err)
	// Test delete picture on a slide which does not contain any pictures
	assert.EqualError(t, f.DeleteImage(257, 999), ErrPictureNotExist{999}.Error())
	// Add same pictures on different slides
	id1, err := f.AddImage(256, filepath.Join("test", "images", "2.png"), nil)
	assert.NoError(t, err)
	id2, err := f.AddImage(258, filepath.Join("test", "images", "2.png"), nil)
	assert.NoError(t, err)
	// Test delete picture on a slide, the image should be preserved
	assert.NoError(t, f.DeleteImage(256, id1))
	_, ok := f.Pkg.Load("ppt/media/image3.png")
	assert.True(t, ok, "The image should be preserved while another slide references it")
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestDeleteImage.pptx")))
	assert.NoError(t, f.Close())

	f, err = OpenFile(filepath.Join("test", "TestDeleteImage.pptx"))
	assert.NoError(t, err)
	// Test delete same picture on different slide, the image should be removed
	assert.NoError(t, f.DeleteImage(258, id2))
	_, ok = f.Pkg.Load("ppt/media/image3.png")
	assert.False(t, ok, "The image should be removed after the last reference")
	assert.NoError(t, f.SaveAs(filepath.Join("test", "TestDeleteImage2.pptx")))

	// Test delete picture on not exists slide
	assert.EqualError(t, f.DeleteImage(999, 1), ErrSlideNotExist{999}.Error())
	// Test delete picture which does not exist
	assert.EqualError(t, f.DeleteImage(258, id2), ErrPictureNotExist{id2}.Error())
	assert.NoError(t, f.Close())

	// Test delete picture on a presentation which does not contain the slide
	assert.EqualError(t, NewFile().DeleteImage(999, 1), ErrSlideNotExist{999}.Error())

	f, err = OpenFile(filepath.Join("test", "presentation.pptx"))
	assert.NoError(t, err)
	// Made two pictures of the slide reference the same image relationship
	pics, err := f.GetPictures(258)
	assert.NoError(t, err)
	assert.Len(t, pics, 2)
	pics[1].BlipFill.Blip.Embed = pics[0].BlipFill.Blip.Embed
	assert.NoError(t, f.DeleteImage(258, pics[0].NonVisualPictureProperties.CommonNonVisualProperties.ID))
	// The image relationship must be kept, the second picture still uses it
	rels, err := f.relsReader(f.getSlideRelsPath("ppt/slides/slide3.xml"))
	assert.NoError(t, err)
	assert.NotNil(t, rels)
	rels.mu.Lock()
	found := false
	for _, rel := range rels.Relationships {
		if rel.ID == "rId2" {
			found = true
		}
	}
	rels.mu.Unlock()
	assert.True(t, found, "The shared image relationship should be kept")
	assert.NoError(t, f.Close())
}

func TestEnsureImageContentType(t *testing.T) {
	f := NewFile()
	assert.NoError(t, f.ensureImageContentType("png", "image/png"))
	assert.NoError(t, f.ensureImageContentType("jpg", "image/jpeg"))
	content, err := f.contentTypesReader()
	assert.NoError(t, err)
	content.mu.Lock()
	found := false
	for _, d := range content.Defaults {
		if d.Extension == "jpg" {
			assert.Equal(t, "image/jpeg", d.ContentType)
			found = true
		}
	}
	content.mu.Unlock()
	assert.True(t, found)
	// Test ensure image content type with unsupported charset content types
	f.ContentTypes = nil
	f.Pkg.Store(defaultXMLPathContentTypes, macCyrillicCharset)
	assert.EqualError(t, f.ensureImageContentType("png", "image/png"), "XML syntax error on line 1: invalid UTF-8")
	assert.NoError(t, f.Close())
}

// prepareTestPresentation provides a function to open the test presentation
// file for test and add images to the slides.
func prepareTestPresentation() (*File, error) {
	f, err := OpenFile(filepath.Join("test", "presentation.pptx"))
	if err != nil {
		return nil, err
	}

	// Test add image to the empty slide with offset
	if _, err = f.AddImage(257, filepath.Join("test", "images", "1.jpg"),
		&PictureOptions{OffsetX: 1333500, OffsetY: 1143000}); err != nil {
		return nil, err
	}

	file, err := os.ReadFile(filepath.Join("test", "images", "1.png"))
	if err != nil {
		return nil, err
	}

	// Test add image to the slide which already contains pictures from bytes
	if _, err = f.AddImageFromBytes(258, file, &PictureOptions{Name: "Logo"}); err != nil {
		return nil, err
	}

	return f, nil
}

// TestSetImage checks the replace of the image data by the slide id and the
// relationship id: the empty data returns the error, the slide that is not
// in the presentation returns the error, the relationship that is not the
// image relationship returns the error, the replaced image data is kept on
// the round trip.
func TestSetImage(t *testing.T) {
	f, err := OpenFile(testFixture)
	assert.NoError(t, err)
	defer f.Close()

	newData, err := os.ReadFile(filepath.Join("test", "images", "2.png"))
	assert.NoError(t, err)

	// the empty image data
	err = f.SetImage(256, "rId2", nil)
	assert.EqualError(t, err, ErrImageEmpty.Error())
	// the slide that is not in the presentation
	err = f.SetImage(999, "rId2", newData)
	assert.EqualError(t, err, ErrSlideNotExist{SlideID: 999}.Error())
	// the relationship that is not in the rels part of the slide
	err = f.SetImage(256, "rId999", newData)
	assert.EqualError(t, err, ErrImageNotExist{RelationshipID: "rId999"}.Error())
	// the relationship that is not the image relationship (the slide layout
	// relationship)
	err = f.SetImage(256, "rId1", newData)
	assert.EqualError(t, err, ErrImageNotExist{RelationshipID: "rId1"}.Error())

	// the image data is replaced (the picture of the slide 256, the
	// relationship rId2)
	oldData, err := f.GetImage(256, "rId2")
	assert.NoError(t, err)
	assert.NotEqual(t, newData, oldData)
	err = f.SetImage(256, "rId2", newData)
	assert.NoError(t, err)
	data, err := f.GetImage(256, "rId2")
	assert.NoError(t, err)
	assert.Equal(t, newData, data)

	// the replaced image data is kept on the round trip
	out := filepath.Join(t.TempDir(), "set_image.pptx")
	assert.NoError(t, f.SaveAs(out))
	reopened, err := OpenFile(out)
	assert.NoError(t, err)
	defer reopened.Close()
	data, err = reopened.GetImage(256, "rId2")
	assert.NoError(t, err)
	assert.Equal(t, newData, data)
}

// TestSVGSize checks the get of the size of the SVG image by the width and
// the height attributes of the root svg element (the pixel values, the in,
// cm, mm, pt and pc units are converted to the pixels at 96 dpi) and by the
// viewBox attribute.
func TestSVGSize(t *testing.T) {
	tests := []struct {
		data          string
		width, height int
		ok            bool
	}{
		{`<svg width="100" height="50"/>`, 100, 50, true},
		{`<svg width="100px" height="50px"/>`, 100, 50, true},
		{`<svg width='1in' height='2in'/>`, 96, 192, true},
		{`<svg width="5cm" height="2.54cm"/>`, 189, 96, true},
		{`<svg width="10mm" height="25.4mm"/>`, 38, 96, true},
		{`<svg width="72pt" height="144pt"/>`, 96, 192, true},
		{`<svg width="1pc" height="2pc"/>`, 16, 32, true},
		{`<svg viewBox="0 0 100 200"/>`, 100, 200, true},
		{`<svg width="10" viewBox="0 0 100 200"/>`, 100, 200, true},
		{`<svg xmlns="http://www.w3.org/2000/svg" width="10" height="20"/>`, 10, 20, true},
		{`<svg/>`, 0, 0, false},
		{`<svg width="10"/>`, 0, 0, false},
		{`<svg viewBox="0 0 100"/>`, 0, 0, false},
		{`<svg viewBox="0 0 -100 200"/>`, 0, 0, false},
		{`<html/>`, 0, 0, false},
	}
	for _, tt := range tests {
		t.Run(tt.data, func(t *testing.T) {
			width, height, ok := svgSize([]byte(tt.data))
			assert.Equal(t, tt.ok, ok)
			if ok {
				assert.Equal(t, tt.width, width)
				assert.Equal(t, tt.height, height)
			}
		})
	}
}

// TestSVGAttrPixels checks the get of the pixel value of the attribute of
// the root svg tag: the value without the unit and with the px, in, cm, mm,
// pt and pc units, the value with the unsupported unit, the absent attribute
// and the invalid value.
func TestSVGAttrPixels(t *testing.T) {
	tag := []byte(`<svg width="10px" height="1in">`)
	if v, ok := svgAttrPixels(tag, "width"); !ok || v != 10 {
		t.Fatalf("expected the width 10 pixels, got %d %v", v, ok)
	}
	if v, ok := svgAttrPixels(tag, "height"); !ok || v != 96 {
		t.Fatalf("expected the height 96 pixels, got %d %v", v, ok)
	}
	// the absent attribute
	if _, ok := svgAttrPixels(tag, "viewBox"); ok {
		t.Fatal("the absent attribute is found")
	}
	// the unsupported unit
	if _, ok := svgAttrPixels([]byte(`<svg width="10em">`), "width"); ok {
		t.Fatal("the value with the unsupported unit is parsed")
	}
	// the invalid value
	if _, ok := svgAttrPixels([]byte(`<svg width="0">`), "width"); ok {
		t.Fatal("the zero value is parsed")
	}
	if _, ok := svgAttrPixels([]byte(`<svg width="abc">`), "width"); ok {
		t.Fatal("the value without the number is parsed")
	}
	// the attribute with the same prefix (the data-width attribute is not
	// the width attribute)
	if _, ok := svgAttrPixels([]byte(`<svg data-width="10">`), "width"); ok {
		t.Fatal("the attribute with the prefix is parsed")
	}
}
