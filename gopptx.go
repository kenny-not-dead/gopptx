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
	"encoding/xml"
	"io"
	"io/fs"
	"os"
	"path/filepath"
	"sync"
)

// File define a populated slides file struct.
type File struct {
	mu            sync.Mutex
	checked       sync.Map
	options       *Options
	tempFiles     sync.Map
	slideMap      map[string]string
	streams       map[string]*StreamWriter
	xmlAttr       sync.Map
	CharsetReader func(charset string, input io.Reader) (rdr io.Reader, err error)
	ContentTypes  *contentTypes
	Path          string
	Pkg           sync.Map
	Presentation  *pptxPresentation
	Relationships sync.Map
	Slide         sync.Map
	SlideCount    int
	Theme         *decodeTheme
	ZipWriter     func(io.Writer) ZipWriter
}

// ZipWriter defines an interface for writing files to a ZIP archive. It
// provides methods to create new files within the archive, add files from a
// filesystem, and close the archive when writing is complete.
type ZipWriter interface {
	Create(name string) (io.Writer, error)
	AddFS(fsys fs.FS) error
	Close() error
}

type Options struct {
	MaxCalcIterations uint
	Password          string
	RawCellValue      bool
	UnzipSizeLimit    int64
	UnzipXMLSizeLimit int64
	TmpDir            string
	ShortDatePattern  string
	LongDatePattern   string
	LongTimePattern   string
}

// OpenFile take the name of a presentation file and returns a populated
// presentation file struct for it.
//
// Close the file by Close function after opening the slides.
func OpenFile(filename string, opts ...Options) (*File, error) {
	file, err := os.Open(filepath.Clean(filename))
	if err != nil {
		return nil, err
	}
	f, err := OpenReader(file, opts...)
	if err != nil {
		if closeErr := file.Close(); closeErr != nil {
			return f, closeErr
		}
		return f, err
	}
	f.Path = filename
	return f, file.Close()
}

// newFile is object builder
func newFile() *File {
	return &File{
		options:   &Options{UnzipSizeLimit: UnzipSizeLimit, UnzipXMLSizeLimit: StreamChunkSize},
		xmlAttr:   sync.Map{},
		checked:   sync.Map{},
		slideMap:  make(map[string]string),
		tempFiles: sync.Map{},
		ZipWriter: func(w io.Writer) ZipWriter { return zip.NewWriter(w) },
	}
}

// checkOpenReaderOptions check and validate options field value for open
// reader.
func (f *File) checkOpenReaderOptions() error {
	if f.options.UnzipSizeLimit == 0 {
		f.options.UnzipSizeLimit = UnzipSizeLimit
		if f.options.UnzipXMLSizeLimit > f.options.UnzipSizeLimit {
			f.options.UnzipSizeLimit = f.options.UnzipXMLSizeLimit
		}
	}
	if f.options.UnzipXMLSizeLimit == 0 {
		f.options.UnzipXMLSizeLimit = StreamChunkSize
		if f.options.UnzipSizeLimit < f.options.UnzipXMLSizeLimit {
			f.options.UnzipXMLSizeLimit = f.options.UnzipSizeLimit
		}
	}
	if f.options.UnzipXMLSizeLimit > f.options.UnzipSizeLimit {
		return ErrOptionsUnzipSizeLimit
	}
	return nil //f.checkDateTimePattern()
}

// OpenReader read data stream from io.Reader and return a populated
// presentation file.
func OpenReader(r io.Reader, opts ...Options) (*File, error) {
	b, err := io.ReadAll(r)
	if err != nil {
		return nil, err
	}
	f := newFile()
	f.options = f.getOptions(opts...)
	if err = f.checkOpenReaderOptions(); err != nil {
		return nil, err
	}
	zr, err := zip.NewReader(bytes.NewReader(b), int64(len(b)))
	if err != nil {
		return nil, err
	}

	file, slideCount, err := f.ReadZipReader(zr)
	if err != nil {
		return nil, err
	}

	f.SlideCount = slideCount
	for k, v := range file {
		f.Pkg.Store(k, v)
	}
	if f.slideMap, err = f.getSlideMap(); err != nil {
		return f, err
	}
	f.Theme, err = f.themeReader()
	return f, err
}

// getOptions provides a function to parse the optional settings for open
// and reading presentation.
func (f *File) getOptions(opts ...Options) *Options {
	options := f.options
	for _, opt := range opts {
		options = &opt
	}
	return options
}

// Creates new XML decoder with charset reader.
func (f *File) xmlNewDecoder(rdr io.Reader) (ret *xml.Decoder) {
	ret = xml.NewDecoder(rdr)
	ret.CharsetReader = f.CharsetReader
	return
}

// slideReader provides a function to get the parsed Slide by slide name.
func (f *File) slideReader(slideName string) (slide *Slide, err error) {
	var (
		path string
		ok   bool
	)
	if err = checkSlideName(slideName); err != nil {
		return
	}
	if path, ok = f.getSlideXMLPath(slideName); !ok {
		err = ErrSlideNotExist{slideName}
		return
	}
	if s, ok := f.Slide.Load(path); ok && s != nil {
		slide = s.(*Slide)
		return
	}

	slide = new(Slide)
	if attrs, ok := f.xmlAttr.Load(path); !ok {
		d := f.xmlNewDecoder(bytes.NewReader(namespaceStrictToTransitional(f.readBytes(path))))
		if attrs == nil {
			attrs = []xml.Attr{}
		}
		attrs = append(attrs.([]xml.Attr), getRootElement(d)...)
		f.xmlAttr.Store(path, attrs)
	}
	if err = f.xmlNewDecoder(bytes.NewReader(namespaceStrictToTransitional(f.readBytes(path)))).
		Decode(slide); err != nil && err != io.EOF {
		return
	}
	err = nil
	if _, ok = f.checked.Load(path); !ok {
		f.checked.Store(path, true)
	}
	f.Slide.Store(path, slide)

	return
}

// checkSlideName check whether there are illegal characters in the slide name.
func checkSlideName(name string) error {
	if name == "" {
		return ErrSlideNameBlank
	}

	return nil
}
