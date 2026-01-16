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
	"fmt"
	"io"
	"path/filepath"
	"strconv"
	"strings"
)

// presentationReader provides a function to get the pointer to the presentation.xml
// structure after deserialization.
func (f *File) presentationReader() (*decodePresentation, error) {
	var err error
	if f.Presentation == nil {
		wbPath := f.getPresentationPath()
		f.Presentation = new(decodePresentation)

		if attrs, ok := f.xmlAttr.Load(wbPath); !ok {
			fmt.Println(attrs)
			d := f.xmlNewDecoder(bytes.NewReader(namespaceStrictToTransitional(f.readXML(wbPath))))
			fmt.Println(d)
			if attrs == nil {
				attrs = []xml.Attr{}
			}
			attrs = append(attrs.([]xml.Attr), getRootElement(d)...)
			f.xmlAttr.Store(wbPath, attrs)
			f.addNameSpaces(wbPath, SourceRelationship)
		}
		if err = f.xmlNewDecoder(bytes.NewReader(namespaceStrictToTransitional(f.readXML(wbPath)))).
			Decode(f.Presentation); err != nil && err != io.EOF {
			return f.Presentation, err
		}
	}

	return f.Presentation, err
}

// getPresentationPath provides a function to get the path of the presentation.xml in
// the presentation.
func (f *File) getPresentationPath() (path string) {
	if rels, _ := f.relsReader(defaultXMLPathRels); rels != nil {
		rels.mu.Lock()
		defer rels.mu.Unlock()
		for _, rel := range rels.Relationships {
			if rel.Type == SourceRelationshipOfficeDocument {
				path = strings.TrimPrefix(rel.Target, "/")
				return
			}
		}
	}
	return
}

// getPresentationRelsPath provides a function to get the path of the presentation.xml.rels
// in the presentation.
func (f *File) getPresentationRelsPath() (path string) {
	wbPath := f.getPresentationPath()
	wbDir := filepath.Dir(wbPath)
	if wbDir == "." {
		path = "_rels/" + filepath.Base(wbPath) + ".rels"
		return
	}
	path = strings.TrimPrefix(filepath.Dir(wbPath)+"/_rels/"+filepath.Base(wbPath)+".rels", "/")
	return
}

// setPresentation update presentation.
func (f *File) setPresentation(slideID, rid int) {
	presentation, _ := f.presentationReader()
	presentation.Slides.Slide = append(presentation.Slides.Slide, decodeSlideID{
		SlideID:        slideID,
		RelationshipID: "rId" + strconv.Itoa(rid),
	})
}
