// Copyright 2026 kenny-not-dead. All rights reserved.
// Use of this source code is governed by a BSD-style license
// that can be found in the LICENSE file.
//
// Package gopptx provides functionality to create and manipulate PowerPoint
// (.pptx) files in Go, using the Office Open XML (ECMA-376) format.

package gopptx

import (
	"encoding/xml"
	"sync"
)

// relationships describe references from parts to other internal resources
// in the package or to external resources.
type relationships struct {
	mu            sync.Mutex
	XMLName       xml.Name       `xml:"http://schemas.openxmlformats.org/package/2006/relationships Relationships"`
	Relationships []relationship `xml:"Relationship"`
}

// relationship contains relations which maps id and XML.
type relationship struct {
	ID         string `xml:"Id,attr"`
	Target     string `xml:",attr"`
	Type       string `xml:",attr"`
	TargetMode string `xml:",attr,omitempty"`
}

// TODO
type presentation struct {
	XMLName                xml.Name          `xml:"http://schemas.openxmlformats.org/presentationml/2006/main p:presentation"`
	AlternateContent       *alternateContent `xml:"mc:AlternateContent"`
	DecodeAlternateContent *innerXML         `xml:"http://schemas.openxmlformats.org/markup-compatibility/2006 AlternateContent"`
	MasterSlide            masterSlideList   `xml:"p:sldMasterIdLst"`
	Slides                 *slideList        `xml:"p:sldIdLst,omitempty"`
	SlideSize              *slideSize        `xml:"p:sldSz,omitempty"`
	NotesSize              *slideSize        `xml:"p:notesSz,omitempty"`
}

// TODO
type masterSlideList struct {
	MasterSlide slideID `xml:"p:sldMasterId"`
}

type slideList struct {
	Slide []slideID `xml:"p:sldId"`
}

// decodeSlideID defines a slide in this presentation. Slide data is stored in a
// separate part.
type slideID struct {
	RelationshipID string `xml:"r:id,attr"`
	SlideID        int    `xml:"id,attr"`
}

// decodePresentation contains elements and attributes that encompass the data
// content of the presentation.
type decodePresentation struct {
	XMLName                xml.Name              `xml:"http://schemas.openxmlformats.org/presentationml/2006/main presentation"`
	AlternateContent       *alternateContent     `xml:"mc:AlternateContent"`
	DecodeAlternateContent *innerXML             `xml:"http://schemas.openxmlformats.org/markup-compatibility/2006 AlternateContent"`
	MasterSlide            decodeMasterSlideList `xml:"sldMasterIdLst"`
	Slides                 *decodeSlideList      `xml:"sldIdLst,omitempty"`
	SlideSize              *slideSize            `xml:"sldSz,omitempty"`
	NotesSize              *slideSize            `xml:"notesSz,omitempty"`
}

type decodeMasterSlideList struct {
	MasterSlide decodeSlideID `xml:"sldMasterId"`
}

type decodeSlideList struct {
	Slide []decodeSlideID `xml:"sldId"`
}

// decodeSlideID defines a slide in this presentation. Slide data is stored in a
// separate part.
type decodeSlideID struct {
	RelationshipID string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
	SlideID        int    `xml:"id,attr"`
}

type slideSize struct {
	CX int `xml:"cx,attr"`
	CY int `xml:"cy,attr"`
}
