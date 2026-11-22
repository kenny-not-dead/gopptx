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
	"sync/atomic"
	"testing"
)

// BenchmarkOpenFile measures the time to open the test presentation and to
// close it again.
func BenchmarkOpenFile(b *testing.B) {
	b.ReportAllocs()
	for b.Loop() {
		f, err := OpenFile(testFixture)
		if err != nil {
			b.Error(err)
		}
		if err := f.Close(); err != nil {
			b.Error(err)
		}
	}
}

// newPresentationWithSlides provides a helper to build a new presentation
// with 10 slides (the helper is used outside the timed loop, the same way
// the excelize library uses the newSheetWithSet helper).
func newPresentationWithSlides() *File {
	f := NewFile()
	for i := 0; i < 10; i++ {
		_, _ = f.NewSlide()
	}
	return f
}

// newPresentationWithShapes provides a helper to build a new presentation
// with 100 shapes with the text body on the first slide.
func newPresentationWithShapes() *File {
	f := NewFile()
	slideID := f.GetSlideList()[0]
	for i := 0; i < 100; i++ {
		_, _ = f.CreateShape(slideID,
			DecodeShapeProperties{
				Xfrm: &DecodeXfrm{
					Offset:  &Offset{X: i * 100000, Y: 1000000},
					Extents: &Extents{CX: 2000000, CY: 1000000},
				},
			},
			DecodeTextBody{
				BodyProperties: &DecodeBodyProperties{},
				ListStyle:      &struct{}{},
				Paragraph:      []DecodeParagraph{{Runs: []DecodeRuns{{Text: "shape " + strconv.Itoa(i)}}}},
			})
	}
	return f
}

// BenchmarkNewSlide measures the time to create a new presentation with 10
// slides.
func BenchmarkNewSlide(b *testing.B) {
	b.RunParallel(func(pb *testing.PB) {
		for pb.Next() {
			newPresentationWithSlides()
		}
	})
}

// BenchmarkCreateShape measures the time to create a new presentation with
// 100 shapes with the text body on the first slide.
func BenchmarkCreateShape(b *testing.B) {
	b.RunParallel(func(pb *testing.PB) {
		for pb.Next() {
			newPresentationWithShapes()
		}
	})
}

// BenchmarkFile_SaveAs measures the time to create a new presentation with
// 100 shapes with the text body and to save it to a file.
func BenchmarkFile_SaveAs(b *testing.B) {
	b.ReportAllocs()
	dir := b.TempDir()
	var counter int64
	b.RunParallel(func(pb *testing.PB) {
		for pb.Next() {
			f := newPresentationWithShapes()
			name := filepath.Join(dir, "save_"+strconv.FormatInt(atomic.AddInt64(&counter, 1), 10)+".pptx")
			if err := f.SaveAs(name); err != nil {
				b.Error(err)
			}
			_ = f.Close()
		}
	})
}

// BenchmarkSetShapeText measures the time to set the text of the shape of
// the test presentation.
func BenchmarkSetShapeText(b *testing.B) {
	b.ReportAllocs()
	f, err := OpenFile(testFixture)
	if err != nil {
		b.Fatal(err)
	}
	defer f.Close()
	const text = "Line 1\nLine 2\nLine 3"
	for b.Loop() {
		if err := f.SetShapeText(257, 8, text); err != nil {
			b.Error(err)
		}
	}
}

// BenchmarkWrite measures the time to build a presentation with 10 slides
// with 50 shapes each and to save it to a file, the same way the
// BenchmarkWrite of the excelize library builds and saves a spreadsheet
// with 200000 cells.
func BenchmarkWrite(b *testing.B) {
	b.ReportAllocs()
	dir := b.TempDir()
	var counter int64
	for b.Loop() {
		f := NewFile()
		for i := 0; i < 9; i++ {
			if _, err := f.NewSlide(); err != nil {
				b.Error(err)
			}
		}
		for _, slideID := range f.GetSlideList() {
			for i := 0; i < 50; i++ {
				if _, err := f.CreateShape(slideID, DecodeShapeProperties{},
					DecodeTextBody{
						BodyProperties: &DecodeBodyProperties{},
						ListStyle:      &struct{}{},
						Paragraph:      []DecodeParagraph{{Runs: []DecodeRuns{{Text: "This is test data"}}}},
					}); err != nil {
					b.Error(err)
				}
			}
		}
		name := filepath.Join(dir, "write_"+strconv.FormatInt(atomic.AddInt64(&counter, 1), 10)+".pptx")
		if err := f.SaveAs(name); err != nil {
			b.Error(err)
		}
		_ = f.Close()
	}
}
