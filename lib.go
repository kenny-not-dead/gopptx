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
	"fmt"
	"io"
	"os"
	"strconv"
	"strings"
)

// ReadZipReader extract presentation with given options.
func (f *File) ReadZipReader(r *zip.Reader) (map[string][]byte, int, error) {
	var (
		err     error
		docPart = map[string]string{
			"[content_types].xml": defaultXMLPathContentTypes,
		}
		fileList  = make(map[string][]byte, len(r.File))
		slides    int
		unzipSize int64
	)
	for _, v := range r.File {
		fileSize := v.FileInfo().Size()
		unzipSize += fileSize
		if unzipSize > f.options.UnzipSizeLimit {
			return fileList, slides, newUnzipSizeLimitError(f.options.UnzipSizeLimit)
		}
		fileName := strings.ReplaceAll(v.Name, "\\", "/")
		if partName, ok := docPart[strings.ToLower(fileName)]; ok {
			fileName = partName
		}
		if strings.HasPrefix(strings.ToLower(fileName), "ppt/slides/slide") {
			slides++
			if fileSize > f.options.UnzipXMLSizeLimit && !v.FileInfo().IsDir() {
				tempFile, err := f.unzipToTemp(v)
				if tempFile != "" {
					f.tempFiles.Store(fileName, tempFile)
				}
				if err == nil {
					continue
				}
			}
		}
		if fileList[fileName], err = readFile(v); err != nil {
			return nil, 0, err
		}
	}
	return fileList, slides, nil
}

// Read file content as string in an archive file.
func readFile(file *zip.File) ([]byte, error) {
	rc, err := file.Open()
	if err != nil {
		return nil, err
	}

	dat := make([]byte, 0, file.FileInfo().Size())
	buff := bytes.NewBuffer(dat)
	_, _ = io.Copy(buff, rc)
	return buff.Bytes(), rc.Close()
}

// unzipToTemp unzip the zip entity to the system temporary directory and
// returned the unzipped file path.
func (f *File) unzipToTemp(zipFile *zip.File) (string, error) {
	tmp, err := os.CreateTemp(f.options.TmpDir, "excelize-")
	if err != nil {
		return "", err
	}
	rc, err := zipFile.Open()
	if err != nil {
		return tmp.Name(), err
	}
	if _, err = io.Copy(tmp, rc); err != nil {
		return tmp.Name(), err
	}
	if err = rc.Close(); err != nil {
		return tmp.Name(), err
	}
	return tmp.Name(), tmp.Close()
}

// namespaceStrictToTransitional provides a method to convert Strict and
// Transitional namespaces.
func namespaceStrictToTransitional(content []byte) []byte {
	namespaceTranslationDic := map[string]string{
		StrictNameSpaceDocumentPropertiesVariantTypes: NameSpaceDocumentPropertiesVariantTypes.Value,
		StrictNameSpaceDrawingMLMain:                  NameSpaceDrawingMLMain,
		StrictNameSpaceExtendedProperties:             NameSpaceExtendedProperties,
		StrictNameSpacePresentationMLMain:             NameSpacePresentationML.Value,
	}
	for s, n := range namespaceTranslationDic {
		content = bytesReplace(content, []byte(s), []byte(n), -1)
	}
	return content
}

// bytesReplace replace source bytes with given target.
func bytesReplace(s, source, target []byte, n int) []byte {
	if n == 0 {
		return s
	}

	if len(source) < len(target) {
		return bytes.Replace(s, source, target, n)
	}

	if n < 0 {
		n = len(s)
	}

	var wid, i, j, w int
	for i, j = 0, 0; i < len(s) && j < n; j++ {
		wid = bytes.Index(s[i:], source)
		if wid < 0 {
			break
		}

		w += copy(s[w:], s[i:i+wid])
		w += copy(s[w:], target)
		i += wid + len(source)
	}

	w += copy(s[w:], s[i:])
	return s[:w]
}

// readXML provides a function to read XML content as bytes.
func (f *File) readXML(name string) []byte {
	if content, _ := f.Pkg.Load(name); content != nil {
		return content.([]byte)
	}
	if content, ok := f.streams[name]; ok {
		return content.rawData.buf.Bytes()
	}
	return []byte{}
}

// saveFileList provides a function to update given file content in file list
// of presentation.
func (f *File) saveFileList(name string, content []byte) {
	f.Pkg.Store(name, append([]byte(xml.Header), content...))
}

// getRootElement extract root element attributes by given XML decoder.
func getRootElement(d *xml.Decoder) []xml.Attr {
	tokenIdx := 0
	for {
		token, _ := d.Token()
		if token == nil {
			break
		}
		switch startElement := token.(type) {
		case xml.StartElement:
			tokenIdx++
			if tokenIdx == 1 {
				var ns bool
				for i := 0; i < len(startElement.Attr); i++ {
					if startElement.Attr[i].Value == NameSpacePresentationML.Value &&
						startElement.Attr[i].Name == NameSpacePresentationML.Name {
						ns = true
					}
				}
				if !ns {
					startElement.Attr = append(startElement.Attr, NameSpacePresentationML)
				}
				return startElement.Attr
			}
		}
	}
	return nil
}

// addNameSpaces provides a function to add an XML attribute by the given
// component part path.
func (f *File) addNameSpaces(path string, ns xml.Attr) {
	exist := false
	mc := false
	ignore := -1
	if attrs, ok := f.xmlAttr.Load(path); ok {
		for i, attr := range attrs.([]xml.Attr) {
			if attr.Name.Local == ns.Name.Local && attr.Name.Space == ns.Name.Space {
				exist = true
			}
			if attr.Name.Local == "Ignorable" && getXMLNamespace(attr.Name.Space, attrs.([]xml.Attr)) == "mc" {
				ignore = i
			}
			if attr.Name.Local == "mc" && attr.Name.Space == "xmlns" {
				mc = true
			}
		}
	}
	if !exist {
		attrs, _ := f.xmlAttr.Load(path)
		if attrs == nil {
			attrs = []xml.Attr{}
		}
		attrs = append(attrs.([]xml.Attr), ns)
		f.xmlAttr.Store(path, attrs)
		if !mc {
			attrs = append(attrs.([]xml.Attr), SourceRelationshipCompatibility)
			f.xmlAttr.Store(path, attrs)
		}
		if ignore == -1 {
			attrs = append(attrs.([]xml.Attr), xml.Attr{
				Name:  xml.Name{Local: "Ignorable", Space: "mc"},
				Value: ns.Name.Local,
			})
			f.xmlAttr.Store(path, attrs)
			return
		}
		f.setIgnorableNameSpace(path, ignore, ns)
	}
}

// getXMLNamespace extract XML namespace from specified element name and attributes.
func getXMLNamespace(space string, attr []xml.Attr) string {
	for _, attribute := range attr {
		if attribute.Value == space {
			return attribute.Name.Local
		}
	}
	return space
}

// inStrSlice provides a method to check if an element is present in an array,
// and return the index of its location, otherwise return -1.
func inStrSlice(a []string, x string, caseSensitive bool) int {
	for idx, n := range a {
		if !caseSensitive && strings.EqualFold(x, n) {
			return idx
		}
		if x == n {
			return idx
		}
	}
	return -1
}

// setIgnorableNameSpace provides a function to set XML namespace as ignorable
// by the given attribute.
func (f *File) setIgnorableNameSpace(path string, index int, ns xml.Attr) {
	ignorableNS := []string{"c14", "cdr14", "a14", "pic14", "x14", "xdr14", "x14ac", "dsp", "mso14", "dgm14", "x15", "x12ac", "x15ac", "xr", "xr2", "xr3", "xr4", "xr5", "xr6", "xr7", "xr8", "xr9", "xr10", "xr11", "xr12", "xr13", "xr14", "xr15", "x15", "x16", "x16r2", "mo", "mx", "mv", "o", "v"}
	xmlAttrs, _ := f.xmlAttr.Load(path)
	if inStrSlice(strings.Fields(xmlAttrs.([]xml.Attr)[index].Value), ns.Name.Local, true) == -1 && inStrSlice(ignorableNS, ns.Name.Local, true) != -1 {
		xmlAttrs.([]xml.Attr)[index].Value = strings.TrimSpace(fmt.Sprintf("%s %s", xmlAttrs.([]xml.Attr)[index].Value, ns.Name.Local))
		f.xmlAttr.Store(path, xmlAttrs)
	}
}

// readBytes read file as bytes by given path.
func (f *File) readBytes(name string) []byte {
	content := f.readXML(name)
	if len(content) != 0 {
		return content
	}
	file, err := f.readTemp(name)
	if err != nil {
		return content
	}
	content, _ = io.ReadAll(file)
	f.Pkg.Store(name, content)
	_ = file.Close()
	return content
}

// readTemp read file from system temporary directory by given path.
func (f *File) readTemp(name string) (file *os.File, err error) {
	path, ok := f.tempFiles.Load(name)
	if !ok {
		return
	}
	file, err = os.Open(path.(string))
	return
}

func (s *decodeSlideID) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	s.SlideID = 0
	s.RelationshipID = ""

	for _, attr := range start.Attr {
		if attr.Name.Local == "id" {
			switch attr.Name.Space {
			case "http://schemas.openxmlformats.org/officeDocument/2006/relationships":
				s.RelationshipID = attr.Value
			case "":
				val, err := strconv.Atoi(attr.Value)
				if err != nil {
					return err
				}

				s.SlideID = val
			default:
				fmt.Printf("❓ Unexpected namespace for 'id': %q\n", attr.Name.Space)
			}
		}
	}

	return d.Skip()
}
