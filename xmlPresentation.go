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
	Type       string `xml:",attr"`
	Target     string `xml:",attr"`
	TargetMode string `xml:",attr,omitempty"`
}

// TODO
type presentation struct {
	XMLName                xml.Name          `xml:"http://schemas.openxmlformats.org/presentationml/2006/main p:presentation"`
	XMLNSA                 string            `xml:"xmlns:a,attr"`
	XMLNSP                 string            `xml:"xmlns:p,attr"`
	XMLNSR                 string            `xml:"xmlns:r,attr"`
	XMLNSP14               string            `xml:"xmlns:p14,attr"`
	XMLNSP15               string            `xml:"xmlns:p15,attr"`
	XMLNSMC                string            `xml:"xmlns:mc,attr"`
	AlternateContent       *alternateContent `xml:"mc:AlternateContent"`
	DecodeAlternateContent *innerXML         `xml:"http://schemas.openxmlformats.org/markup-compatibility/2006 AlternateContent"`
	MasterSlide            masterSlideList   `xml:"p:sldMasterIdLst"`
	NotesMaster            *notesMasterList  `xml:"p:notesMasterIdLst,omitempty"`
	Slides                 *slideList        `xml:"p:sldIdLst,omitempty"`
	SlideSize              *slideSize        `xml:"p:sldSz,omitempty"`
	NotesSize              *notesSize        `xml:"p:notesSz,omitempty"`
	DefaultTextStyle       *innerXML         `xml:"p:defaultTextStyle"`
	ExtList                *innerXML         `xml:"p:extLst"`
}

// TODO
type masterSlideList struct {
	MasterSlide slideID `xml:"p:sldMasterId"`
}

// notesMasterList defines the notes master id list (p:notesMasterIdLst)
// element of the presentation.
type notesMasterList struct {
	NotesMaster notesMasterID `xml:"p:notesMasterId"`
}

// notesMasterID defines a notes master in this presentation.
type notesMasterID struct {
	RelationshipID string `xml:"r:id,attr"`
}

// TODO
type slideList struct {
	Slide []slideID `xml:"p:sldId"`
}

// TODO
type slideID struct {
	SlideID        int    `xml:"id,attr"`
	RelationshipID string `xml:"r:id,attr"`
}

// decodePresentation contains elements and attributes that encompass the data
// content of the presentation.
type decodePresentation struct {
	XMLName                xml.Name              `xml:"http://schemas.openxmlformats.org/presentationml/2006/main presentation"`
	AlternateContent       *alternateContent     `xml:"mc:AlternateContent"`
	DecodeAlternateContent *innerXML             `xml:"http://schemas.openxmlformats.org/markup-compatibility/2006 AlternateContent"`
	MasterSlide            decodeMasterSlideList `xml:"sldMasterIdLst"`
	NotesMaster            decodeNotesMasterList `xml:"notesMasterIdLst,omitempty"`
	Slides                 *decodeSlideList      `xml:"sldIdLst,omitempty"`
	SlideSize              *slideSize            `xml:"sldSz,omitempty"`
	NotesSize              *notesSize            `xml:"notesSz,omitempty"`
	DefaultTextStyle       *innerXML             `xml:"defaultTextStyle"`
	ExtList                *innerXML             `xml:"extLst"`
}

type decodeMasterSlideList struct {
	MasterSlide decodeSlideID `xml:"sldMasterId"`
}

// decodeNotesMasterList defines the notes master id list (notesMasterIdLst)
// element of the presentation.
type decodeNotesMasterList struct {
	NotesMaster decodeNotesMasterID `xml:"notesMasterId"`
}

// decodeNotesMasterID defines a notes master in this presentation.
type decodeNotesMasterID struct {
	RelationshipID string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
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
	CY   int     `xml:"cy,attr"`
	CX   int     `xml:"cx,attr"`
	Type *string `xml:"type,attr,omitempty"`
}

type notesSize struct {
	CX int `xml:"cx,attr"`
	CY int `xml:"cy,attr"`
}
