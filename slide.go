// Copyright 2026 kenny-not-dead. All rights reserved.
// Use of this source code is governed by a BSD-style license
// that can be found in the LICENSE file.
//
// Package gopptx provides functionality to create and manipulate PowerPoint
// (.pptx) files in Go, using the Office Open XML (ECMA-376) format.

package gopptx

import (
	"bytes"
	"fmt"
	"io"
	"path/filepath"
	"strings"
)

// contentTypesReader provides a function to get the pointer to the
// [Content_Types].xml structure after deserialization.
func (f *File) contentTypesReader() (*contentTypes, error) {
	if f.ContentTypes == nil {
		f.ContentTypes = new(contentTypes)
		f.ContentTypes.mu.Lock()
		defer f.ContentTypes.mu.Unlock()
		if err := f.xmlNewDecoder(bytes.NewReader(namespaceStrictToTransitional(f.readXML(defaultXMLPathContentTypes)))).
			Decode(f.ContentTypes); err != nil && err != io.EOF {
			return f.ContentTypes, err
		}
	}
	return f.ContentTypes, nil
}

// getSlideMap provides a function to get slide name and XML file path map
// of the presentation.
func (f *File) getSlideMap() (map[string]string, error) {
	maps := map[string]string{}
	wb, err := f.presentationReader()
	if err != nil {
		return nil, err
	}

	rels, err := f.relsReader(f.getPresentationRelsPath())
	if err != nil {
		return nil, err
	}
	if rels == nil {
		return maps, nil
	}
	for _, slide := range wb.Slides.Slide {
		for _, rel := range rels.Relationships {
			if rel.ID == slide.ID {
				slideXMLPath := f.getSlidePath(rel.Target)
				if _, ok := f.Pkg.Load(slideXMLPath); ok {
					maps[slide.ID] = slideXMLPath
				}
				if _, ok := f.tempFiles.Load(slideXMLPath); ok {
					maps[slide.ID] = slideXMLPath
				}
			}
		}
	}

	return maps, nil
}

// relsReader provides a function to get the pointer to the structure
// after deserialization of relationships parts.
func (f *File) relsReader(path string) (*relationships, error) {
	rels, _ := f.Relationships.Load(path)
	if rels == nil {
		if _, ok := f.Pkg.Load(path); ok {
			c := relationships{}
			if err := f.xmlNewDecoder(bytes.NewReader(namespaceStrictToTransitional(f.readXML(path)))).
				Decode(&c); err != nil && err != io.EOF {
				return nil, err
			}
			f.Relationships.Store(path, &c)
		}
	}
	if rels, _ = f.Relationships.Load(path); rels != nil {
		return rels.(*relationships), nil
	}
	return nil, nil
}

// getSlidePath construct a target XML as ppt/slides/slide%d by split
// path, compatible with different types of relative paths in
// presentation.xml.rels, for example: slides/slide%d.xml
// and /ppt/slides/slide%d.xml
func (f *File) getSlidePath(relTarget string) (path string) {
	path = filepath.ToSlash(strings.TrimPrefix(
		strings.ReplaceAll(filepath.Clean(fmt.Sprintf("%s/%s", filepath.Dir(f.getPresentationPath()), relTarget)), "\\", "/"), "/"))
	if strings.HasPrefix(relTarget, "/") {
		path = filepath.ToSlash(strings.TrimPrefix(strings.ReplaceAll(filepath.Clean(relTarget), "\\", "/"), "/"))
	}
	return path
}

// getSlideXMLPath provides a function to get XML file path by given slide name.
func (f *File) getSlideXMLPath(slide string) (string, bool) {
	var (
		name string
		ok   bool
	)
	for slideName, filePath := range f.slideMap {
		if strings.EqualFold(slideName, slide) {
			name, ok = filePath, true
			break
		}
	}
	return name, ok
}

// GetShapes provides a function to get shapes by given slide id.
func (f *File) GetShapes(slideID string) ([]Shape, error) {
	var shapes []Shape
	ws, err := f.slideReader(slideID)
	if err != nil {
		return shapes, err
	}

	return ws.getShapes(), err
}

// getShapes returns shapes of the slide.
func (ws *Slide) getShapes() []Shape {
	return ws.CommonSlideData.ShapeTree.Shape
}

// GetGroupShapeProperties provides a function to get group shape properties by given slide id.
func (f *File) GetGroupShapeProperties(slideID string) (*GroupShapeProperties, error) {
	s, err := f.slideReader(slideID)
	if err != nil {
		return nil, err
	}

	return s.getGroupShapeProperties(), err
}

// getGroupShapeProperties returns group shape properties of the slide.
func (ws *Slide) getGroupShapeProperties() *GroupShapeProperties {
	return ws.CommonSlideData.ShapeTree.GroupShapeProperties
}

// GetNonVisualGroupShapeProperties provides a function to get non visual group shape properties by given slide id.
func (f *File) GetNonVisualGroupShapeProperties(slideID string) (*NonVisualGroupShapeProperties, error) {
	s, err := f.slideReader(slideID)
	if err != nil {
		return nil, err
	}

	return s.getNonVisualGroupShapeProperties(), err
}

// NonVisualGroupShapeProperties returns non visual group shape properties of the slide.
func (ws *Slide) getNonVisualGroupShapeProperties() *NonVisualGroupShapeProperties {
	return ws.CommonSlideData.ShapeTree.NonVisualGroupShapeProperties
}
