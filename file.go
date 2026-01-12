// Copyright 2026 kenny-not-dead. All rights reserved.
// Use of this source code is governed by a BSD-style license
// that can be found in the LICENSE file.
//
// Package gopptx provides functionality to create and manipulate PowerPoint
// (.pptx) files in Go, using the Office Open XML (ECMA-376) format.

package gopptx

import (
	"bytes"
	_ "embed"
	"encoding/xml"
	"fmt"
	"io"
	"math"
	"os"
	"path/filepath"
	"sort"
	"strings"
	"sync"
)

// NewFile provides a function to create new file by default template.
// For example:
//
//	f := NewFile()
func NewFile(opts ...Options) *File {
	f := newFile()
	f.Pkg.Store(defaultXMLPathRels, []byte(xml.Header+templateRels))
	f.Pkg.Store(defaultXMLPathDocPropsApp, []byte(xml.Header+templateDocpropsApp))
	f.Pkg.Store(defaultXMLPathDocPropsCore, []byte(xml.Header+templateDocpropsCore))
	f.Pkg.Store(defaultXMLPathPresentationRels, []byte(xml.Header+templatePresentationRels))
	f.Pkg.Store(defaultXMLPathTheme, []byte(xml.Header+templateTheme))
	f.Pkg.Store(defaultXMLPathSlideMaster, []byte(xml.Header+templateSlideMaster))
	f.Pkg.Store(defaultXMLPathSlideMasterRels, []byte(xml.Header+templateSlideMasterRels))
	f.Pkg.Store(defaultXMLPathSlideLayout, []byte(xml.Header+templateSlideLayout))
	f.Pkg.Store(defaultXMLPathSlideLayoutRels, []byte(xml.Header+templateSlideLayoutRels))
	f.Pkg.Store(defaultXMLPathSlide, []byte(xml.Header+TemplateSlide))
	f.Pkg.Store(defaultXMLPathSlideRels, []byte(xml.Header+TemplateSlideRels))
	f.Pkg.Store(defaultXMLPathPresentation, []byte(xml.Header+templatePresentation))
	f.Pkg.Store(defaultXMLPathPresProps, []byte(xml.Header+templatePresProps))
	f.Pkg.Store(defaultXMLPathContentTypes, []byte(xml.Header+templateContentTypes))
	f.SlideCount = 1
	f.ContentTypes, _ = f.contentTypesReader()
	f.Presentation, _ = f.presentationReader()
	f.Relationships = sync.Map{}
	// rels, _ := f.relsReader(defaultXMLPathPresentationRels)
	// f.Relationships.Store(defaultXMLPathPresentationRels, rels)
	// f.slideMap[defaultXMLSlideID] = defaultXMLPathSlide
	// // ws, err := f.slideReader(defaultXMLSlideID)
	// if err != nil {
	// 	fmt.Println(err)
	// }
	// f.Slide.Store(defaultXMLPathSlide, ws)
	f.Theme, _ = f.themeReader()
	f.options = f.getOptions(opts...)
	return f
}

// Save provides a function to override the spreadsheet with origin path.
func (f *File) Save(opts ...Options) error {
	if f.Path == "" {
		return ErrSave
	}
	for i := range opts {
		f.options = &opts[i]
	}
	return f.SaveAs(f.Path, *f.options)
}

// SaveAs provides a function to create or update to a presentation at the
// provided path.
func (f *File) SaveAs(name string, opts ...Options) error {
	if len(name) > MaxFilePathLength {
		return ErrMaxFilePathLength
	}
	f.Path = name
	if _, ok := supportedContentTypes[strings.ToLower(filepath.Ext(f.Path))]; !ok {
		return ErrPresentationFileFormat
	}
	file, err := os.OpenFile(filepath.Clean(name), os.O_WRONLY|os.O_TRUNC|os.O_CREATE, os.ModePerm)
	if err != nil {
		return err
	}
	defer file.Close()
	return f.Write(file, opts...)
}

// Close closes and cleanup the open temporary file for the spreadsheet.
func (f *File) Close() error {
	var firstErr error

	for _, stream := range f.streams {
		_ = stream.rawData.Close()
	}
	f.streams = nil
	f.tempFiles.Range(func(k, v interface{}) bool {
		if err := os.Remove(v.(string)); err != nil && firstErr == nil {
			firstErr = err
		}
		return true
	})
	f.tempFiles.Clear()

	return firstErr
}

// Write provides a function to write to an io.Writer.
func (f *File) Write(w io.Writer, opts ...Options) error {
	_, err := f.WriteTo(w, opts...)
	return err
}

// WriteTo implements io.WriterTo to write the file.
func (f *File) WriteTo(w io.Writer, opts ...Options) (int64, error) {
	for i := range opts {
		f.options = &opts[i]
	}
	if len(f.Path) != 0 {
		contentType, ok := supportedContentTypes[strings.ToLower(filepath.Ext(f.Path))]
		if !ok {
			return 0, ErrPresentationFileFormat
		}
		if err := f.setContentTypePartProjectExtensions(contentType); err != nil {
			return 0, err
		}
	}
	buf, err := f.WriteToBuffer()
	if err != nil {
		return 0, err
	}
	return buf.WriteTo(w)
}

// WriteToBuffer provides a function to get bytes.Buffer from the saved file,
// and it allocates space in memory. Be careful when the file size is large.
func (f *File) WriteToBuffer() (*bytes.Buffer, error) {
	buf := new(bytes.Buffer)
	zw := f.ZipWriter(buf)

	if err := f.writeToZip(zw); err != nil {
		_ = zw.Close()
		return buf, err
	}
	if err := zw.Close(); err != nil {
		return buf, err
	}
	// f.writeZip64LFH(buf)

	return buf, nil
}

// writeToZip provides a function to write to ZipWriter.
func (f *File) writeToZip(zw ZipWriter) error {
	f.contentTypesWriter()
	// f.presentationWriter()
	// TODO: MasterWritter
	// TODO: MasterLayoutWritter
	f.slideWriter() // TODO: check wrire slide data
	f.relsWriter()
	f.themeWriter()

	for path, stream := range f.streams {
		fi, err := zw.Create(path)
		if err != nil {
			return err
		}
		var from io.Reader
		if from, err = stream.rawData.Reader(); err != nil {
			_ = stream.rawData.Close()
			return err
		}
		written, err := io.Copy(fi, from)
		if err != nil {
			return err
		}
		if written > math.MaxUint32 {
			f.zip64Entries = append(f.zip64Entries, path)
		}
	}
	var (
		n                int
		err              error
		files, tempFiles []string
	)
	f.Pkg.Range(func(path, content interface{}) bool {
		if _, ok := f.streams[path.(string)]; ok {
			return true
		}
		files = append(files, path.(string))
		return true
	})
	sort.Sort(sort.Reverse(sort.StringSlice(files)))
	for _, path := range files {
		var fi io.Writer
		if fi, err = zw.Create(path); err != nil {
			break
		}
		content, _ := f.Pkg.Load(path)
		if n, err = fi.Write(content.([]byte)); int64(n) > math.MaxUint32 {
			f.zip64Entries = append(f.zip64Entries, path)
		}
	}
	f.tempFiles.Range(func(path, content interface{}) bool {
		if _, ok := f.Pkg.Load(path); ok {
			return true
		}
		tempFiles = append(tempFiles, path.(string))
		return true
	})
	sort.Sort(sort.Reverse(sort.StringSlice(tempFiles)))
	for _, path := range tempFiles {
		var fi io.Writer
		if fi, err = zw.Create(path); err != nil {
			break
		}
		if n, err = fi.Write(f.readBytes(path)); int64(n) > math.MaxUint32 {
			f.zip64Entries = append(f.zip64Entries, path)
		}
	}
	return err
}

// setContentTypePartProjectExtensions provides a function to set the content
// type for relationship parts and the main document part.
func (f *File) setContentTypePartProjectExtensions(contentType string) error {
	content, err := f.contentTypesReader()
	if err != nil {
		return err
	}
	content.mu.Lock()
	defer content.mu.Unlock()

	for idx, o := range content.Overrides {
		if o.PartName == "/"+defaultXMLPathPresentation {
			content.Overrides[idx].ContentType = contentType
		}
	}

	return err
}

// contentTypesWriter provides a function to save [Content_Types].xml after
// serialize structure.
func (f *File) contentTypesWriter() {
	if f.ContentTypes != nil {
		output, _ := xml.Marshal(f.ContentTypes)
		f.saveFileList(defaultXMLPathContentTypes, output)
	}
}

// presentationWriter provides a function to save presentation.xml after serialize
// structure.
func (f *File) presentationWriter() {
	if f.Presentation != nil {
		if f.Presentation.DecodeAlternateContent != nil {
			f.Presentation.AlternateContent = &alternateContent{
				Content: f.Presentation.DecodeAlternateContent.Content,
				XMLNSMC: SourceRelationshipCompatibility.Value,
			}
		}
		f.Presentation.DecodeAlternateContent = nil
		output, _ := xml.Marshal(f.Presentation)
		f.saveFileList(f.getPresentationPath(), replaceRelationshipsBytes(f.replaceNameSpaceBytes(f.getPresentationPath(), output)))
	}
}

// replaceRelationshipsBytes; Some tools that read spreadsheet files have very
// strict requirements about the structure of the input XML. This function is
// a horrible hack to fix that after the XML marshalling is completed.
func replaceRelationshipsBytes(content []byte) []byte {
	sourceXmlns := []byte(`xmlns:relationships="http://schemas.openxmlformats.org/officeDocument/2006/relationships" relationships`)
	targetXmlns := []byte("r")
	return bytesReplace(content, sourceXmlns, targetXmlns, -1)
}

// replaceNameSpaceBytes provides a function to replace the XML root element
// attribute by the given component part path and XML content.
func (f *File) replaceNameSpaceBytes(path string, content []byte) []byte {
	// sourceXmlns := []byte(`xmlns="http://schemas.openxmlformats.org/presentationml/2006/main">`)
	// targetXmlns := []byte(templateNamespaceIDMap)
	// if attrs, ok := f.xmlAttr.Load(path); ok {
	// 	targetXmlns = []byte(genXMLNamespace(attrs.([]xml.Attr)))
	// }
	// return bytesReplace(content, sourceXmlns, bytes.ReplaceAll(targetXmlns, []byte(" mc:Ignorable=\"r\""), []byte{}), -1)

	// if !(strings.HasPrefix(path, "ppt/slides/slide") && strings.HasSuffix(path, ".xml")) &&
	// 	path != "ppt/presentation.xml" {
	// 	return content
	// }

	// if strings.HasPrefix(path, "ppt/slides/slide") {
	// 	xmlns := ` xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="p14" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"`
	// 	content = bytes.Replace(content, []byte("<p:sld>"), []byte("<p:sld"+xmlns+">"), 1)
	// }

	// if path == "ppt/presentation.xml" {
	// 	xmlns := ` xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="p14" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"`
	// 	content = bytes.Replace(content, []byte("<p:presentation>"), []byte("<p:presentation"+xmlns+">"), 1)
	// }

	return content
}

// genXMLNamespace generate serialized XML attributes with a multi namespace
// by given element attributes.
func genXMLNamespace(attr []xml.Attr) string {
	var rootElement string
	for _, v := range attr {
		if lastSpace := getXMLNamespace(v.Name.Space, attr); lastSpace != "" {
			if lastSpace == NameSpaceXML {
				lastSpace = "xml"
			}
			rootElement += fmt.Sprintf("%s:%s=\"%s\" ", lastSpace, v.Name.Local, v.Value)
			continue
		}
		rootElement += fmt.Sprintf("%s=\"%s\" ", v.Name.Local, v.Value)
	}
	return strings.TrimSpace(rootElement) + ">"
}

// themeWriter provides a function to save xl/theme/theme1.xml after serialize
// structure.
func (f *File) themeWriter() {
	newColor := func(c *decodeComplexTypeColor) complexTypeColor {
		return complexTypeColor{
			ScrgbColor:  c.ScrgbColor,
			SrgbColor:   c.SrgbColor,
			HslColor:    c.HslColor,
			SystemColor: c.SystemColor,
			SchemeColor: c.SchemeColor,
			PresetColor: c.PresetColor,
		}
	}
	newFontScheme := func(c *decodeFontCollection) fontCollection {
		return fontCollection{
			Latin: c.Latin,
			Ea:    c.Ea,
			Cs:    c.Cs,
			Font:  c.Font,
		}
	}
	if f.Theme != nil {
		output, _ := xml.Marshal(theme{
			XMLNSa: NameSpaceDrawingML.Value,
			XMLNSr: SourceRelationship.Value,
			Name:   f.Theme.Name,
			ThemeElements: baseStyles{
				ColorScheme: colorScheme{
					Name:     f.Theme.ThemeElements.ColorScheme.Name,
					Dk1:      newColor(&f.Theme.ThemeElements.ColorScheme.Dk1),
					Lt1:      newColor(&f.Theme.ThemeElements.ColorScheme.Lt1),
					Dk2:      newColor(&f.Theme.ThemeElements.ColorScheme.Dk2),
					Lt2:      newColor(&f.Theme.ThemeElements.ColorScheme.Lt2),
					Accent1:  newColor(&f.Theme.ThemeElements.ColorScheme.Accent1),
					Accent2:  newColor(&f.Theme.ThemeElements.ColorScheme.Accent2),
					Accent3:  newColor(&f.Theme.ThemeElements.ColorScheme.Accent3),
					Accent4:  newColor(&f.Theme.ThemeElements.ColorScheme.Accent4),
					Accent5:  newColor(&f.Theme.ThemeElements.ColorScheme.Accent5),
					Accent6:  newColor(&f.Theme.ThemeElements.ColorScheme.Accent6),
					Hlink:    newColor(&f.Theme.ThemeElements.ColorScheme.Hlink),
					FolHlink: newColor(&f.Theme.ThemeElements.ColorScheme.FolHlink),
				},
				FontScheme: fontScheme{
					Name:      f.Theme.ThemeElements.FontScheme.Name,
					MajorFont: newFontScheme(&f.Theme.ThemeElements.FontScheme.MajorFont),
					MinorFont: newFontScheme(&f.Theme.ThemeElements.FontScheme.MinorFont),
				},
				FormatScheme: styleMatrix{
					Name:            f.Theme.ThemeElements.FormatScheme.Name,
					FillStyleList:   f.Theme.ThemeElements.FormatScheme.FillStyleList,
					LineStyleList:   f.Theme.ThemeElements.FormatScheme.LineStyleList,
					EffectStyleList: f.Theme.ThemeElements.FormatScheme.EffectStyleList,
					BgFillStyleList: f.Theme.ThemeElements.FormatScheme.BgFillStyleList,
				},
			},
		})
		f.saveFileList(defaultXMLPathTheme, f.replaceNameSpaceBytes(defaultXMLPathTheme, output))
	}
}

// relsWriter provides a function to save relationships after
// serialize structure.
func (f *File) relsWriter() {
	f.Relationships.Range(func(path, rel interface{}) bool {
		if rel != nil {
			output, _ := xml.Marshal(rel.(*relationships))
			if strings.HasPrefix(path.(string), "ppt/slides/_rels/slide") {
				output = f.replaceNameSpaceBytes(path.(string), output)
			}
			f.saveFileList(path.(string), replaceRelationshipsBytes(output))
		}
		return true
	})
}

// slideWriter provides a function to save xl/worksheets/sheet%d.xml after
// serialize structure.
func (f *File) slideWriter() {
	var (
		arr     []byte
		buffer  = bytes.NewBuffer(arr)
		encoder = xml.NewEncoder(buffer)
	)
	f.Slide.Range(func(p, ws interface{}) bool {
		if ws != nil {
			slide := ws.(*decodeSlide)

			// f.addNameSpaces(p.(string), SourceRelationship)

			if slide.DecodeAlternateContent != nil {
				slide.AlternateContent = &alternateContent{
					Content: slide.DecodeAlternateContent.Content,
					XMLNSMC: SourceRelationshipCompatibility.Value,
				}
			}
			slide.DecodeAlternateContent = nil
			// reusing buffer
			_ = encoder.Encode(slide)
			f.saveFileList(p.(string), replaceRelationshipsBytes(f.replaceNameSpaceBytes(p.(string), buffer.Bytes())))
			_, ok := f.checked.Load(p.(string))
			if ok {
				f.Slide.Delete(p.(string))
				f.checked.Delete(p.(string))
			}
			buffer.Reset()
		}
		return true
	})
}
