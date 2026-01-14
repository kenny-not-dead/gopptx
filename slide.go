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

// NewSlide provides the function to create a new slide and
// returns the index of the slide in the presentation after it appended.
func (f *File) NewSlide() (int, error) {
	presentation, err := f.presentationReader()
	if err != nil {
		return -1, err
	}

	f.SlideCount++

	slideID := defaultXMLSlideID
	for _, s := range presentation.Slides.Slide {
		if s.SlideID >= slideID {
			slideID = s.SlideID + 1
		}
	}

	nextFileIndex := len(presentation.Slides.Slide) + 1
	fileName := "slide" + strconv.Itoa(nextFileIndex)

	// Update [Content_Types].xml
	_ = f.setContentTypes("/ppt/slides/_rels/"+fileName+".xml.rels", ContentTypeRelationships)
	_ = f.setContentTypes("/ppt/slides/"+fileName+".xml", ContentTypeSlideML)

	// Update presentation.xml.rels
	rID := f.addRels(f.getPresentationRelsPath(), SourceRelationshipSlide, fmt.Sprintf("slides/%s.xml", fileName), "")

	// Create new sheet /ppt/slides/sheet%d.xml
	_ = f.setSlide(nextFileIndex, slideID)

	// Update presentation.xml
	f.setPresentation(slideID, rID)

	return slideID, nil
}

// setContentTypes provides a function to read and update property of contents
// type of the spreadsheet.
func (f *File) setContentTypes(partName, contentType string) error {
	content, err := f.contentTypesReader()
	if err != nil {
		return err
	}
	content.mu.Lock()
	defer content.mu.Unlock()
	content.Overrides = append(content.Overrides, contentTypeOverride{
		PartName:    partName,
		ContentType: contentType,
	})
	return err
}

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

// setSlide provides a function to update slide property by given index.
func (f *File) setSlide(index int, id int) error {
	slideXMLPath := "ppt/slides/slide" + strconv.Itoa(index) + ".xml"
	f.slideMap[id] = slideXMLPath

	slide := decodeSlide{}
	err := xml.Unmarshal([]byte(xml.Header+TemplateSlide), &slide)
	if err != nil {
		return err
	}

	f.Slide.Store(slideXMLPath, &slide)
	f.xmlAttr.Store(slideXMLPath, []xml.Attr{NameSpacePresentationML}) // TODO: check attr

	return nil
}

// getSlideMap provides a function to get slide name and XML file path map
// of the presentation.
func (f *File) getSlideMap() (map[int]string, error) {
	maps := map[int]string{}
	presentation, err := f.presentationReader()
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
	for _, slide := range presentation.Slides.Slide {
		for _, rel := range rels.Relationships {
			if rel.ID == slide.RelationshipID {
				slideXMLPath := f.getSlidePath(rel.Target)
				if _, ok := f.Pkg.Load(slideXMLPath); ok {
					maps[slide.SlideID] = slideXMLPath
				}
				if _, ok := f.tempFiles.Load(slideXMLPath); ok {
					maps[slide.SlideID] = slideXMLPath
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

// getSlideXMLPath provides a function to get XML file path by given slide id.
func (f *File) getSlideXMLPath(id int) (string, bool) {
	path, ok := f.slideMap[id]
	return path, ok
}

// GetShapes provides a function to get shapes by given slide id.
func (f *File) GetShapes(slideID int) ([]decodeShape, error) {
	var shapes []decodeShape
	ws, err := f.slideReader(slideID)
	if err != nil {
		return shapes, err
	}

	return ws.getShapes(), err
}

// getShapes returns shapes of the slide.
func (ws *decodeSlide) getShapes() []decodeShape {
	return ws.CommonSlideData.ShapeTree.Shape
}

// GetGroupShapeProperties provides a function to get group shape properties by given slide id.
func (f *File) GetGroupShapeProperties(slideID int) (*decodeGroupShapeProperties, error) {
	s, err := f.slideReader(slideID)
	if err != nil {
		return nil, err
	}

	return s.getGroupShapeProperties(), err
}

// getGroupShapeProperties returns group shape properties of the slide.
func (ws *decodeSlide) getGroupShapeProperties() *decodeGroupShapeProperties {
	return ws.CommonSlideData.ShapeTree.GroupShapeProperties
}

// GetNonVisualGroupShapeProperties provides a function to get non visual group shape properties by given slide id.
func (f *File) GetNonVisualGroupShapeProperties(slideID int) (*decodeNonVisualGroupShapeProperties, error) {
	s, err := f.slideReader(slideID)
	if err != nil {
		return nil, err
	}

	return s.getNonVisualGroupShapeProperties(), err
}

// NonVisualGroupShapeProperties returns non visual group shape properties of the slide.
func (ws *decodeSlide) getNonVisualGroupShapeProperties() *decodeNonVisualGroupShapeProperties {
	return ws.CommonSlideData.ShapeTree.NonVisualGroupShapeProperties
}
