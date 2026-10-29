// Copyright 2026 kenny-not-dead. All rights reserved.
// Use of this source code is governed by a BSD-style license
// that can be found in the LICENSE file.
//
// Package gopptx provides functionality to create and manipulate PowerPoint
// (.pptx) files in Go, using the Office Open XML (ECMA-376) format.

package gopptx

import (
	"encoding/xml"
)

// decodeNotesSlide directly maps the notes slide part (p:notes).
type decodeNotesSlide struct {
	XMLName                xml.Name              `xml:"notes"`
	CommonSlideData        decodeNotesSlideData  `xml:"cSld"`
	ColorMapOverride       *decodeClrMapOverride `xml:"clrMapOvr,omitempty"`
	AlternateContent       *alternateContent     `xml:"mc:AlternateContent"`
	DecodeAlternateContent *innerXML             `xml:"http://schemas.openxmlformats.org/markup-compatibility/2006 AlternateContent"`
}

// decodeNotesSlideData directly maps the common slide data (cSld) element of
// the notes slide.
type decodeNotesSlideData struct {
	ShapeTree decodeNotesShapeTree `xml:"spTree"`
}

// decodeNotesShapeTree directly maps the shape tree (spTree) element of the
// notes slide. Only the placeholder shapes (p:sp) are modeled, in the
// document order: the standard notes slide contains the slide image
// placeholder (sldImg), the notes body placeholder (body) and the slide
// number placeholder (sldNum).
type decodeNotesShapeTree struct {
	NonVisualGroupShapeProperties *decodeNonVisualGroupShapeProperties `xml:"nvGrpSpPr"`
	GroupShapeProperties          *decodeGroupShapeProperties          `xml:"grpSpPr"`
	Shape                         []decodeShape                        `xml:"sp"`
}

// notesBodyShape returns the shape of the notes body placeholder (the
// placeholder with the ph type="body") of the notes slide, or nil when the
// notes slide has no body placeholder.
func (ds *decodeNotesSlide) notesBodyShape() *decodeShape {
	for i := range ds.CommonSlideData.ShapeTree.Shape {
		shape := &ds.CommonSlideData.ShapeTree.Shape[i]
		if shape.NonVisualShapeProperties != nil &&
			shape.NonVisualShapeProperties.NonVisualProperties != nil &&
			shape.NonVisualShapeProperties.NonVisualProperties.Ph != nil &&
			shape.NonVisualShapeProperties.NonVisualProperties.Ph.Type != nil &&
			*shape.NonVisualShapeProperties.NonVisualProperties.Ph.Type == "body" {
			return shape
		}
	}
	return nil
}

// NotesSlide defines the notes slide part (p:notes).
type NotesSlide struct {
	XMLName          xml.Name          `xml:"p:notes"`
	XMLNSA           string            `xml:"xmlns:a,attr"`
	XMLNSP           string            `xml:"xmlns:p,attr"`
	XMLNSR           string            `xml:"xmlns:r,attr"`
	CommonSlideData  NotesSlideData    `xml:"p:cSld"`
	ColorMapOverride *ClrMapOverride   `xml:"p:clrMapOvr,omitempty"`
	AlternateContent *alternateContent `xml:"mc:AlternateContent"`
}

// NotesSlideData defines the common slide data (p:cSld) element of the notes
// slide.
type NotesSlideData struct {
	ShapeTree NotesShapeTree `xml:"p:spTree"`
}

// NotesShapeTree defines the shape tree (p:spTree) element of the notes
// slide.
type NotesShapeTree struct {
	NonVisualGroupShapeProperties *NonVisualGroupShapeProperties `xml:"p:nvGrpSpPr"`
	GroupShapeProperties          *GroupShapeProperties          `xml:"p:grpSpPr"`
	Shape                         []Shape                        `xml:"p:sp"`
}
