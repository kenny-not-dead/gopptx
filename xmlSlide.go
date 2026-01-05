// Copyright 2026 kenny-not-dead. All rights reserved.
// Use of this source code is governed by a BSD-style license
// that can be found in the LICENSE file.
//
// Package gopptx provides functionality to create and manipulate PowerPoint
// (.pptx) files in Go, using the Office Open XML (ECMA-376) format.

package gopptx

import (
	"sync"
)

type Slide struct {
	mu              sync.Mutex
	CommonSlideData SlideData `xml:"cSld"`
}

type SlideData struct {
	ShapeTree *ShapeTree `xml:"spTree"`
}

type ShapeTree struct {
	NonVisualGroupShapeProperties *NonVisualGroupShapeProperties `xml:"nvGrpSpPr,omitempty"`
	GroupShapeProperties          *GroupShapeProperties          `xml:"grpSpPr,omitempty"`
	Shape                         []Shape                        `xml:"sp"`
}

type NonVisualGroupShapeProperties struct {
	CommonNonVisualProperties           *CommonNonVisualProperties           `xml:"cNvPr"`
	CommonNonVisualGroupShapeProperties *CommonNonVisualGroupShapeProperties `xml:"cNvGrpSpPr"`
	NonVisualProperties                 *NonVisualProperties                 `xml:"nvPr"`
}

type CommonNonVisualProperties struct {
	ID   int    `xml:"id,attr"`
	Name string `xml:"name,attr"`
}

type CommonNonVisualGroupShapeProperties struct{}

type NonVisualProperties struct{}

type GroupShapeProperties struct {
	Xfrm *Xfrm `xml:"xfrm"`
}

type Xfrm struct {
	Offset       *Offset  `xml:"off"`
	Extents      *Extents `xml:"ext"`
	ChildOffset  *Offset  `xml:"chOff"`
	ChildExtents *Extents `xml:"chExt"`
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
	NonVisualShapeProperties *NonVisualShapeProperties `xml:"nvSpPr"`
	ShapeProperties          *ShapeProperties          `xml:"spPr"`
	TextBody                 *TextBody                 `xml:"txBody,omitempty"`
}

type NonVisualShapeProperties struct {
	CommonNonVisualProperties      *CommonNonVisualProperties      `xml:"cNvPr"`
	CommonNonVisualShapeProperties *CommonNonVisualShapeProperties `xml:"cNvSpPr"`
	NonVisualProperties            *NonVisualProperties            `xml:"nvPr"`
}

type CommonNonVisualShapeProperties struct {
	TxBox      *bool `xml:"txBox,attr,omitempty"`
	ShapeLocks *int  `xml:"spLocks,attr,omitempty"`
}

type ShapeProperties struct {
	Xfrm           *Xfrm           `xml:"xfrm"`
	PresetGeometry *PresetGeometry `xml:"prstGeom,omitempty"`
	NoFill         *any            `xml:"noFill,omitempty"`
	Ln             *Line           `xml:"ln,omitempty"`
}

type PresetGeometry struct {
	Preset          string `xml:"prst,attr"`
	AdjustValueList *any   `xml:"avLst"`
}

type AdjustValue struct {
	Name    string `xml:"name,attr"`
	Formula string `xml:"fmla,attr"`
}

type Line struct {
	Width  int  `xml:"w,attr,omitempty"`
	NoFill *any `xml:"noFill,omitempty"`
}

type TextBody struct {
	BodyProperties *BodyProperties `xml:"bodyPr"`
	Paragraph      []Paragraph     `xml:"p"`
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
	ParagraphProperties       *ParagraphProperties `xml:"pPr,omitempty"`
	Runs                      []Runs               `xml:"r"`
	EndParagraphRunProperties *Runs                `xml:"endParaRPr,omitempty"`
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
