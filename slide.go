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

	// Create new slide /ppt/slides/slide%d.xml and slide rels /ppt/slides/_rels/slide%d.xml.rels
	_ = f.setSlide(nextFileIndex, slideID)

	// Update presentation.xml
	f.setPresentation(slideID, rID)

	return slideID, nil
}

// setContentTypes provides a function to read and update property of contents
// type of the presentation.
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

	relsSlideXMLPath := "ppt/slides/_rels/slide" + strconv.Itoa(index) + ".xml.rels"
	f.Pkg.Store(relsSlideXMLPath, []byte(xml.Header+TemplateSlideRels))

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
	slide, err := f.slideReader(slideID)
	if err != nil {
		return shapes, err
	}

	return slide.getShapes(), err
}

// getShapes returns shapes of the slide.
func (ds *decodeSlide) getShapes() []decodeShape {
	return ds.CommonSlideData.ShapeTree.Shape
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
func (ds *decodeSlide) getGroupShapeProperties() *decodeGroupShapeProperties {
	return ds.CommonSlideData.ShapeTree.GroupShapeProperties
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
func (ds *decodeSlide) getNonVisualGroupShapeProperties() *decodeNonVisualGroupShapeProperties {
	return ds.CommonSlideData.ShapeTree.NonVisualGroupShapeProperties
}

// DeleteSlide provides a function to delete slide in a presentation by given slide id.
func (f *File) DeleteSlide(slideID int) error {
	if idx, _ := f.GetSlideIndex(slideID); f.SlideCount == 1 || idx == -1 {
		return nil
	}

	presentation, _ := f.presentationReader()
	presentationRels, _ := f.relsReader(f.getPresentationRelsPath())

	for idx, v := range presentation.Slides.Slide {
		if v.SlideID != slideID {
			continue
		}

		presentation.Slides.Slide = append(presentation.Slides.Slide[:idx], presentation.Slides.Slide[idx+1:]...)
		var slideXML, rels string
		if presentationRels != nil {
			for _, rel := range presentationRels.Relationships {
				if rel.ID == v.RelationshipID {
					slideXML = f.getSlidePath(rel.Target)
					rels = "ppt/slides/_rels/" + strings.TrimPrefix(slideXML, "ppt/slides/") + ".rels"
				}
			}
		}

		target := f.deleteSlideFromPresentationRels(v.RelationshipID)
		dir := filepath.Dir(target)
		base := filepath.Base(target)

		// Update [Content_Types].xml
		_ = f.removeContentTypesPart(ContentTypeSlideML, target)
		_ = f.removeContentTypesPart(ContentTypeRelationships, filepath.Join(dir, "_rels", base+".rels"))

		delete(f.slideMap, v.SlideID)
		f.Pkg.Delete(slideXML)
		f.Pkg.Delete(rels)
		f.Relationships.Delete(rels)
		f.Slide.Delete(slideXML)
		f.xmlAttr.Delete(slideXML)
		f.SlideCount--
	}

	// TODO: setActiveSlide
	//index, err := f.GetSlideIndex(f.getActiveSlideID())
	//f.SetActiveSlide(index)
	return nil
}

// deleteSlideFromPresentationRels provides a function to remove slide
// relationships by given relationships ID in the file presentation.xml.rels.
func (f *File) deleteSlideFromPresentationRels(rID string) string {
	rels, _ := f.relsReader(f.getPresentationRelsPath())
	rels.mu.Lock()
	defer rels.mu.Unlock()
	for k, v := range rels.Relationships {
		if v.ID == rID {
			rels.Relationships = append(rels.Relationships[:k], rels.Relationships[k+1:]...)
			return v.Target
		}
	}
	return ""
}

// GetSlideIndex provides a function to get a slide index of the presentation by
// the given slide id. If slide doesn't exist, it will return an integer type value -1.
func (f *File) GetSlideIndex(slideID int) (int, error) {
	for index, id := range f.GetSlideList() {
		if id ==  slideID {
			return index, nil
		}
	}
	return -1, nil
}

// GetSlideList provides a function to get slides of the presentation.
func (f *File) GetSlideList() (list []int) {
	presentation, _ := f.presentationReader()
	if presentation != nil {
		for _, slide := range presentation.Slides.Slide {
			list = append(list, slide.SlideID)
		}
	}
	return
}

// GetActiveSlideIndex provides a function to get active slide index of the
// presentation. If not found the active slide will be return integer 0.
func (f *File) GetActiveSlideIndex() (index int) {
	slideID := f.getActiveSlideID()
	presentation, _ := f.presentationReader()
	if presentation != nil {
		for idx, slide := range presentation.Slides.Slide {
			if slide.SlideID == slideID {
				index = idx
				return
			}
		}
	}
	return
}

// getActiveSlideID provides a function to get active slide ID of the
// presentation. If not found the active slide will be return integer 0.
func (f *File) getActiveSlideID() int {
	presentation, _ := f.presentationReader()
	if presentation != nil {
		// TODO: get active slide
		if len(presentation.Slides.Slide) >= 1 {
			return presentation.Slides.Slide[0].SlideID
		}
	}
	return 0
}