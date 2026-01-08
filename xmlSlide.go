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

type Slide struct {
	mu                     sync.Mutex
	XMLName                xml.Name          `xml:"p:sld"`
	CommonSlideData        SlideData         `xml:"p:cSld"`
	AlternateContent       *alternateContent `xml:"mc:AlternateContent"`
	DecodeAlternateContent *innerXML         `xml:"http://schemas.openxmlformats.org/markup-compatibility/2006 AlternateContent"`
}

type SlideData struct {
	ShapeTree ShapeTree `xml:"p:spTree"`
}

type ShapeTree struct {
	NonVisualGroupShapeProperties *NonVisualGroupShapeProperties `xml:"p:nvGrpSpPr,omitempty"`
	GroupShapeProperties          *GroupShapeProperties          `xml:"p:grpSpPr,omitempty"`
	Shape                         []Shape                        `xml:"p:sp"`
}

type NonVisualGroupShapeProperties struct {
	CommonNonVisualProperties           *CommonNonVisualProperties           `xml:"p:cNvPr"`
	CommonNonVisualGroupShapeProperties *CommonNonVisualGroupShapeProperties `xml:"p:cNvGrpSpPr"`
	NonVisualProperties                 *NonVisualProperties                 `xml:"p:nvPr"`
}

type CommonNonVisualProperties struct {
	ID   int    `xml:"id,attr"`
	Name string `xml:"name,attr"`
}

type CommonNonVisualGroupShapeProperties struct{}

type NonVisualProperties struct{}

type GroupShapeProperties struct {
	Xfrm *Xfrm `xml:"a:xfrm"`
}

type Xfrm struct {
	Offset       *Offset  `xml:"a:off"`
	Extents      *Extents `xml:"a:ext"`
	ChildOffset  *Offset  `xml:"a:chOff"`
	ChildExtents *Extents `xml:"a:chExt"`
}

type Offset struct {
	X int `xml:"x,attr"`
	Y int `xml:"y,attr"`
}

type Extents struct {
	CX int `xml:"cx,attr"`
	CY int `xml:"cy,attr"`
}

type Shape struct {
	NonVisualShapeProperties *NonVisualShapeProperties `xml:"p:nvSpPr"`
	ShapeProperties          *ShapeProperties          `xml:"p:spPr"`
	TextBody                 *TextBody                 `xml:"p:txBody,omitempty"`
}

type NonVisualShapeProperties struct {
	CommonNonVisualProperties      *CommonNonVisualProperties      `xml:"p:cNvPr"`
	CommonNonVisualShapeProperties *CommonNonVisualShapeProperties `xml:"p:cNvSpPr"`
	NonVisualProperties            *NonVisualProperties            `xml:"p:nvPr"`
}

type CommonNonVisualShapeProperties struct {
	TxBox      *bool `xml:"txBox,attr,omitempty"`
	ShapeLocks *int  `xml:"spLocks,attr,omitempty"`
}

type ShapeProperties struct {
	Xfrm           *Xfrm           `xml:"a:xfrm"`
	PresetGeometry *PresetGeometry `xml:"a:prstGeom,omitempty"`
	NoFill         *any            `xml:"a:noFill,omitempty"`
	Ln             *Line           `xml:"a:ln,omitempty"`
}

type PresetGeometry struct {
	Preset          string `xml:"prst,attr"`
	AdjustValueList *any   `xml:"a:avLst"`
}

type AdjustValue struct {
	Name    string `xml:"name,attr"`
	Formula string `xml:"fmla,attr"`
}

type Line struct {
	Width  int  `xml:"w,attr,omitempty"`
	NoFill *any `xml:"a:noFill,omitempty"`
}

type TextBody struct {
	BodyProperties *BodyProperties `xml:"a:bodyPr"`
	Paragraph      []Paragraph     `xml:"a:p"`
}

type BodyProperties struct {
	LIns      int    `xml:"lIns,attr,omitempty"`
	RIns      int    `xml:"rIns,attr,omitempty"`
	TIns      int    `xml:"tIns,attr,omitempty"`
	BIns      int    `xml:"bIns,attr,omitempty"`
	Anchor    string `xml:"anchor,attr,omitempty"`
	NoAutofit *any   `xml:"noAutofit,omitempty"`
}

type Paragraph struct {
	ParagraphProperties       *ParagraphProperties `xml:"a:pPr,omitempty"`
	Runs                      []Runs               `xml:"r"`
	EndParagraphRunProperties *Runs                `xml:"a:endParaRPr,omitempty"`
}

type ParagraphProperties struct {
	Indent      int          `xml:"indent,attr,omitempty"`
	Align       *string      `xml:"algn,attr,omitempty"`
	LineSpacing *LineSpacing `xml:"lnSpc,omitempty"`
	BuNone      *struct{}    `xml:"buNone,omitempty"`
}

type LineSpacing struct {
	SpacingPercent *SpacingPercent `xml:"spcPct"`
}

type SpacingPercent struct {
	Val int `xml:"val,attr"`
}

type Runs struct {
	RunProperties *RunProperties `xml:"rPr,omitempty"`
	Text          string         `xml:"t"`
}

type RunProperties struct {
	Bold      *bool      `xml:"b,attr,omitempty"`
	Lang      string     `xml:"lang,attr,omitempty"`
	Size      int        `xml:"sz,attr,omitempty"`
	Space     int        `xml:"spc,attr,omitempty"`
	Strike    string     `xml:"strike,attr,omitempty"`
	SolidFill *SolidFill `xml:"solidFill,omitempty"`
	Latin     *Latin     `xml:"latin,omitempty"`
}

type SolidFill struct {
	SolidRGBColor *SolidRGBColor `xml:"srgbClr"`
}

type SolidRGBColor struct {
	Val string `xml:"val,attr"`
}

type Latin struct {
	Typeface string `xml:"typeface,attr"`
}
