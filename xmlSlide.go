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
	XMLNSA                 string            `xml:"xmlns:a,attr"`
	XMLNSP                 string            `xml:"xmlns:p,attr"`
	XMLNSR                 string            `xml:"xmlns:r,attr"`
	XMLNSP14               string            `xml:"xmlns:p14,attr"`
	XMLNSP15               string            `xml:"xmlns:p15,attr"`
	XMLNSMC                string            `xml:"xmlns:mc,attr"`
	CommonSlideData        SlideData         `xml:"p:cSld"`
	ColorMapOverride       *ClrMapOverride   `xml:"p:clrMapOvr,omitempty"`
	AlternateContent       *alternateContent `xml:"mc:AlternateContent"`
	DecodeAlternateContent *innerXML         `xml:"http://schemas.openxmlformats.org/markup-compatibility/2006 AlternateContent"`
}

type SlideData struct {
	Background *bg       `xml:"p:bg,omitempty"`
	ShapeTree  ShapeTree `xml:"p:spTree"`
}

// bg defines the slide background (p:bg) element of the common slide data.
type bg struct {
	BGProperties *bgProperties `xml:"p:bgPr,omitempty"`
}

// bgProperties defines the background properties (p:bgPr) element.
type bgProperties struct {
	NoFill     *noFill    `xml:"a:noFill,omitempty"`
	SolidFill  *SolidFill `xml:"a:solidFill,omitempty"`
	EffectList *struct{}  `xml:"a:effectLst,omitempty"`
}

// shapeTreeChildKind is the kind of a child of the slide shape tree.
type shapeTreeChildKind int

const (
	shapeTreeChildSp shapeTreeChildKind = iota
	shapeTreeChildPic
	shapeTreeChildCxnSp
)

// shapeTreeOrderEntry is one entry of the document order of the shape tree
// children (the z-order of the shapes): the kind of the child and the index
// of the element in the corresponding slice of the shape tree.
type shapeTreeOrderEntry struct {
	kind  shapeTreeChildKind
	index int
}

type ShapeTree struct {
	NonVisualGroupShapeProperties *NonVisualGroupShapeProperties `xml:"p:nvGrpSpPr,omitempty"`
	GroupShapeProperties          *GroupShapeProperties          `xml:"p:grpSpPr,omitempty"`
	Shape                         []Shape                        `xml:"p:sp"`
	Picture                       []Picture                      `xml:"p:pic"`
	Connector                     []Connector                    `xml:"p:cxnSp"`
	// shapeOrder is the document order of the children (sp/pic/cxnSp) of
	// the shape tree, the z-order of the shapes; when it is empty (the
	// shape tree is built without the document order) the children are
	// emitted as blocks in the struct field order
	shapeOrder []shapeTreeOrderEntry
}

// MarshalXML writes the shape tree, keeping the document order of the
// shapes, so the z-order of a modified slide is preserved on save. The
// element names are written with the literal prefixes, the same way the
// struct tags of the slide do it.
func (st ShapeTree) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = xml.Name{Local: "p:spTree"}
	if err := e.EncodeToken(start); err != nil {
		return err
	}

	if st.NonVisualGroupShapeProperties != nil {
		if err := e.EncodeElement(st.NonVisualGroupShapeProperties, xml.StartElement{
			Name: xml.Name{Local: "p:nvGrpSpPr"},
		}); err != nil {
			return err
		}
	}
	if st.GroupShapeProperties != nil {
		if err := e.EncodeElement(st.GroupShapeProperties, xml.StartElement{
			Name: xml.Name{Local: "p:grpSpPr"},
		}); err != nil {
			return err
		}
	}

	emit := func(name string, v interface{}) error {
		return e.EncodeElement(v, xml.StartElement{Name: xml.Name{Local: name}})
	}

	if len(st.shapeOrder) == 0 {
		// the document order is not recorded: the children are emitted as
		// blocks in the struct field order, the same as before the order
		// was introduced
		for i := range st.Shape {
			if err := emit("p:sp", &st.Shape[i]); err != nil {
				return err
			}
		}
		for i := range st.Picture {
			if err := emit("p:pic", &st.Picture[i]); err != nil {
				return err
			}
		}
		for i := range st.Connector {
			if err := emit("p:cxnSp", &st.Connector[i]); err != nil {
				return err
			}
		}
	} else {
		for i := range st.shapeOrder {
			entry := &st.shapeOrder[i]
			switch entry.kind {
			case shapeTreeChildSp:
				if err := emit("p:sp", &st.Shape[entry.index]); err != nil {
					return err
				}
			case shapeTreeChildPic:
				if err := emit("p:pic", &st.Picture[entry.index]); err != nil {
					return err
				}
			case shapeTreeChildCxnSp:
				if err := emit("p:cxnSp", &st.Connector[entry.index]); err != nil {
					return err
				}
			}
		}
	}

	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

type NonVisualGroupShapeProperties struct {
	CommonNonVisualProperties           *CommonNonVisualProperties           `xml:"p:cNvPr"`
	CommonNonVisualGroupShapeProperties *CommonNonVisualGroupShapeProperties `xml:"p:cNvGrpSpPr"`
	NonVisualProperties                 *NonVisualProperties                 `xml:"p:nvPr"`
}

type GroupShapeProperties struct {
	Xfrm *Xfrm `xml:"a:xfrm"`
}

type Xfrm struct {
	Offset       *Offset  `xml:"a:off"`
	Extents      *Extents `xml:"a:ext"`
	ChildOffset  *Offset  `xml:"a:chOff"`
	ChildExtents *Extents `xml:"a:chExt"`
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

type NonVisualProperties struct {
	Ph *Ph `xml:"p:ph,omitempty"`
}

// Picture defines a picture shape (p:pic) in the slide shape tree.
type Picture struct {
	NonVisualPictureProperties *NonVisualPictureProperties `xml:"p:nvPicPr"`
	BlipFill                   *BlipFill                   `xml:"p:blipFill"`
	ShapeProperties            *ShapeProperties            `xml:"p:spPr"`
}

// NonVisualPictureProperties defines the non-visual properties of a picture
// (p:nvPicPr).
type NonVisualPictureProperties struct {
	CommonNonVisualProperties        *CommonNonVisualProperties        `xml:"p:cNvPr"`
	CommonNonVisualPictureProperties *CommonNonVisualPictureProperties `xml:"p:cNvPicPr"`
	NonVisualProperties              *NonVisualProperties              `xml:"p:nvPr"`
}

// CommonNonVisualPictureProperties defines the common non-visual properties
// of a picture (p:cNvPicPr).
type CommonNonVisualPictureProperties struct {
	PreferRelativeResize *int `xml:"preferRelativeResize,attr,omitempty"`
}

// Connector defines a connector shape (p:cxnSp) in the slide shape tree.
type Connector struct {
	NonVisualConnectorProperties *NonVisualConnectorProperties `xml:"p:nvCxnSpPr"`
	ShapeProperties              *ShapeProperties              `xml:"p:spPr"`
}

// NonVisualConnectorProperties defines the non-visual properties of a
// connector (p:nvCxnSpPr).
type NonVisualConnectorProperties struct {
	CommonNonVisualProperties          *CommonNonVisualProperties          `xml:"p:cNvPr"`
	CommonNonVisualConnectorProperties *CommonNonVisualConnectorProperties `xml:"p:cNvCxnSpPr"`
	NonVisualProperties                *NonVisualProperties                `xml:"p:nvPr"`
}

// CommonNonVisualConnectorProperties defines the common non-visual properties
// of a connector (p:cNvCxnSpPr).
type CommonNonVisualConnectorProperties struct{}

// BlipFill defines the blip fill (p:blipFill) element of a picture.
type BlipFill struct {
	Blip    *Blip    `xml:"a:blip"`
	Stretch *Stretch `xml:"a:stretch"`
}

// Blip defines the blip (a:blip) element of the blip fill. The image part is
// referenced by the relationship id in the embed attribute.
type Blip struct {
	Embed       string       `xml:"r:embed,attr"`
	AlphaModFix *AlphaModFix `xml:"a:alphaModFix,omitempty"`
}

// AlphaModFix defines the alpha modulation fix (a:alphaModFix) element of
// the blip.
type AlphaModFix struct {
	Amount *int `xml:"amt,attr,omitempty"`
}

// Stretch defines the stretch (a:stretch) element of the blip fill.
type Stretch struct {
	FillRect *FillRect `xml:"a:fillRect"`
}

// FillRect defines the fill rectangle (a:fillRect) element of the stretch.
type FillRect struct {
	Left  *int `xml:"l,attr,omitempty"`
	Top   *int `xml:"t,attr,omitempty"`
	Right *int `xml:"r,attr,omitempty"`
	Bot   *int `xml:"b,attr,omitempty"`
}

// ClrMapOverride defines the color map override (p:clrMapOvr) element of
// the slide.
type ClrMapOverride struct {
	MasterClrMapping *MasterClrMapping `xml:"a:masterClrMapping"`
}

// MasterClrMapping defines the master color mapping (a:masterClrMapping)
// element.
type MasterClrMapping struct{}

type CommonNonVisualShapeProperties struct {
	TxBox      *XMLBool    `xml:"txBox,attr,omitempty"`
	ShapeLocks *ShapeLocks `xml:"a:spLocks,omitempty"`
}

type ShapeProperties struct {
	Xfrm           *Xfrm           `xml:"a:xfrm"`
	CustGeom       *innerXML       `xml:"a:custGeom,omitempty"`
	PresetGeometry *PresetGeometry `xml:"a:prstGeom,omitempty"`
	NoFill         *noFill         `xml:"a:noFill,omitempty"`
	SolidFill      *SolidFill      `xml:"a:solidFill,omitempty"`
	Ln             *Line           `xml:"a:ln,omitempty"`
	EffectList     *innerXML       `xml:"a:effectLst,omitempty"`
}

type PresetGeometry struct {
	Preset          string           `xml:"prst,attr"`
	AdjustValueList *AdjustValueList `xml:"a:avLst"`
}

type Line struct {
	Width     *int       `xml:"w,attr,omitempty"`
	NoFill    *noFill    `xml:"a:noFill,omitempty"`
	SolidFill *SolidFill `xml:"a:solidFill,omitempty"`
	PrstDash  *PrstDash  `xml:"a:prstDash,omitempty"`
}

// PrstDash defines the preset dash style (a:prstDash) element of the line.
type PrstDash struct {
	Val string `xml:"val,attr"`
}

type TextBody struct {
	BodyProperties *BodyProperties `xml:"a:bodyPr"`
	ListStyle      *struct{}       `xml:"a:lstStyle,omitempty"`
	Paragraph      []Paragraph     `xml:"a:p"`
}

type BodyProperties struct {
	AnchorCtr        *int         `xml:"anchorCtr,attr,omitempty"`
	Anchor           *string      `xml:"anchor,attr,omitempty"`
	BIns             *int         `xml:"bIns,attr,omitempty"`
	LIns             *int         `xml:"lIns,attr,omitempty"`
	SpcFirstLastPara *int         `xml:"spcFirstLastPara,attr,omitempty"`
	RIns             *int         `xml:"rIns,attr,omitempty"`
	Wrap             *string      `xml:"wrap,attr,omitempty"`
	TIns             *int         `xml:"tIns,attr,omitempty"`
	NormAutofit      *NormAutofit `xml:"a:normAutofit,omitempty"`
	NoAutofit        *NoAutofit   `xml:"a:noAutofit,omitempty"`
}

// NormAutofit defines the normal autofit (a:normAutofit) element of the body
// properties.
type NormAutofit struct {
	FontScale            *int `xml:"fontScale,attr,omitempty"`
	LineSpacingReduction *int `xml:"lnSpcReduction,attr,omitempty"`
}

type Paragraph struct {
	ParagraphProperties       *ParagraphProperties `xml:"a:pPr,omitempty"`
	Runs                      []Runs               `xml:"a:r"`
	EndParagraphRunProperties *RunProperties       `xml:"a:endParaRPr,omitempty"`
}

// ParagraphProperties defines the paragraph properties (a:pPr) element of
// the paragraph.
type ParagraphProperties struct {
	Indent           *int           `xml:"indent,attr,omitempty"`
	Lvl              *int           `xml:"lvl,attr,omitempty"`
	MarL             *int           `xml:"marL,attr,omitempty"`
	RTL              *int           `xml:"rtl,attr,omitempty"`
	Align            *string        `xml:"algn,attr,omitempty"`
	LineSpacing      *LineSpacing   `xml:"a:lnSpc,omitempty"`
	SpacingBefore    *SpacingChange `xml:"a:spcBef,omitempty"`
	SpacingAfter     *SpacingChange `xml:"a:spcAft,omitempty"`
	BuFont           *BulletFont    `xml:"a:buFont,omitempty"`
	BuChar           *BulletChar    `xml:"a:buChar,omitempty"`
	BuNone           *struct{}      `xml:"a:buNone,omitempty"`
	DefRunProperties *RunProperties `xml:"a:defRPr,omitempty"`
}

// BulletFont defines the a:buFont element of the paragraph properties: the
// typeface of the bullet character.
type BulletFont struct {
	Typeface string `xml:"typeface,attr,omitempty"`
}

// BulletChar defines the a:buChar element of the paragraph properties: the
// character used as the bullet.
type BulletChar struct {
	Char string `xml:"char,attr,omitempty"`
}

// SpacingChange defines the spacing before (a:spcBef) or spacing after
// (a:spcAft) element of the paragraph properties.
type SpacingChange struct {
	SpacingPercent *SpacingPercent `xml:"a:spcPct,omitempty"`
	SpacingPts     *SpacingPts     `xml:"a:spcPts,omitempty"`
}

type Runs struct {
	RunProperties *RunProperties `xml:"a:rPr,omitempty"`
	Text          string         `xml:"a:t"`
}

type RunProperties struct {
	Bold           *int       `xml:"b,attr,omitempty"`
	Lang           string     `xml:"lang,attr,omitempty"`
	Size           *int       `xml:"sz,attr,omitempty"`
	Space          *int       `xml:"spc,attr,omitempty"`
	Strike         string     `xml:"strike,attr,omitempty"`
	SolidFill      *SolidFill `xml:"a:solidFill,omitempty"`
	Latin          *Font      `xml:"a:latin,omitempty"`
	EastAsian      *Font      `xml:"a:ea,omitempty"`
	ComplexScript  *Font      `xml:"a:cs,omitempty"`
	HlinkClick     *Hyperlink `xml:"a:hlinkClick,omitempty"`
	HlinkMouseOver *Hyperlink `xml:"a:hlinkMouseOver,omitempty"`
}
type SolidFill struct {
	SolidRGBColor *SolidRGBColor `xml:"a:srgbClr"`
}

type decodeSlide struct {
	mu                     sync.Mutex
	XMLName                xml.Name              `xml:"sld"`
	CommonSlideData        decodeSlideData       `xml:"cSld"`
	ColorMapOverride       *decodeClrMapOverride `xml:"clrMapOvr,omitempty"`
	AlternateContent       *alternateContent     `xml:"mc:AlternateContent"`
	DecodeAlternateContent *innerXML             `xml:"http://schemas.openxmlformats.org/markup-compatibility/2006 AlternateContent"`
}

// decodeClrMapOverride directly maps the clrMapOvr element of the slide.
type decodeClrMapOverride struct {
	MasterClrMapping *MasterClrMapping `xml:"masterClrMapping,omitempty"`
}

type decodeSlideData struct {
	Background *DecodeBackground `xml:"bg,omitempty"`
	ShapeTree  decodeShapeTree   `xml:"spTree"`
}

// DecodeBackground directly maps the bg element of the slide.
type DecodeBackground struct {
	BGProperties *DecodeBGProperties `xml:"bgPr,omitempty"`
}

// DecodeBGProperties directly maps the bgPr element of the slide background.
type DecodeBGProperties struct {
	NoFill     *noFill          `xml:"noFill,omitempty"`
	SolidFill  *DecodeSolidFill `xml:"solidFill,omitempty"`
	EffectList *struct{}        `xml:"effectLst,omitempty"`
}

// decodeShapeTreeChild is a single node of the shape tree in the document
// order. Exactly one of the fields is set.
type decodeShapeTreeChild struct {
	Shape     *decodeShape
	Picture   *decodePicture
	Connector *decodeConnector
}

type decodeShapeTree struct {
	NonVisualGroupShapeProperties *decodeNonVisualGroupShapeProperties `xml:"nvGrpSpPr,omitempty"`
	GroupShapeProperties          *decodeGroupShapeProperties          `xml:"grpSpPr,omitempty"`
	Children                      []decodeShapeTreeChild
}

// UnmarshalXML reads the shape tree, keeping the document order of the
// shapes, so the z-order of the slide is preserved when it is modified and
// saved.
func (dst *decodeShapeTree) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		token, err := d.Token()
		if err != nil {
			return err
		}

		switch t := token.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "nvGrpSpPr":
				nv := new(decodeNonVisualGroupShapeProperties)
				if err := d.DecodeElement(nv, &t); err != nil {
					return err
				}
				dst.NonVisualGroupShapeProperties = nv
			case "grpSpPr":
				gs := new(decodeGroupShapeProperties)
				if err := d.DecodeElement(gs, &t); err != nil {
					return err
				}
				dst.GroupShapeProperties = gs
			case "sp":
				s := new(decodeShape)
				if err := d.DecodeElement(s, &t); err != nil {
					return err
				}
				dst.Children = append(dst.Children, decodeShapeTreeChild{Shape: s})
			case "pic":
				p := new(decodePicture)
				if err := d.DecodeElement(p, &t); err != nil {
					return err
				}
				dst.Children = append(dst.Children, decodeShapeTreeChild{Picture: p})
			case "cxnSp":
				c := new(decodeConnector)
				if err := d.DecodeElement(c, &t); err != nil {
					return err
				}
				dst.Children = append(dst.Children, decodeShapeTreeChild{Connector: c})
			default:
				// unknown child of the shape tree (grpSp, graphicFrame, ...):
				// it is not modeled and is not preserved, the same as before
				if err := d.Skip(); err != nil {
					return err
				}
			}
		case xml.EndElement:
			return nil
		}
	}
}

type decodeNonVisualGroupShapeProperties struct {
	CommonNonVisualProperties           *decodeCommonNonVisualProperties     `xml:"cNvPr"`
	CommonNonVisualGroupShapeProperties *CommonNonVisualGroupShapeProperties `xml:"cNvGrpSpPr"`
	NonVisualProperties                 *decodeNonVisualProperties           `xml:"nvPr"`
}

type CommonNonVisualProperties struct {
	ID             int        `xml:"id,attr"`
	Name           string     `xml:"name,attr"`
	Title          string     `xml:"title,attr,omitempty"`
	HlinkClick     *Hyperlink `xml:"a:hlinkClick,omitempty"`
	HlinkMouseOver *Hyperlink `xml:"a:hlinkMouseOver,omitempty"`
	ExtLst         *ExtLst    `xml:"a:extLst,omitempty"`
}

// ExtLst defines the extension list (a:extLst) element of the common
// non-visual properties. Its content (for example the a16:creationId
// extension) is preserved as-is on the round trip.
type ExtLst struct {
	Content string `xml:",innerxml"`
}

type CommonNonVisualGroupShapeProperties struct{}

type decodeNonVisualProperties struct {
	Ph *Ph `xml:"ph,omitempty"`
}

type Ph struct {
	Idx  *int    `xml:"idx,attr,omitempty"`
	Type *string `xml:"type,attr,omitempty"`
}

type decodeGroupShapeProperties struct {
	Xfrm *DecodeXfrm `xml:"xfrm"`
}

type DecodeXfrm struct {
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

type decodeShape struct {
	NonVisualShapeProperties *decodeNonVisualShapeProperties `xml:"nvSpPr"`
	ShapeProperties          *DecodeShapeProperties          `xml:"spPr"`
	TextBody                 *DecodeTextBody                 `xml:"txBody,omitempty"`
}

// decodePicture directly maps the pic element in the shape tree.
type decodePicture struct {
	NonVisualPictureProperties *decodeNonVisualPictureProperties `xml:"nvPicPr"`
	BlipFill                   *decodeBlipFill                   `xml:"blipFill"`
	ShapeProperties            *DecodeShapeProperties            `xml:"spPr"`
}

// decodeNonVisualPictureProperties directly maps the nvPicPr element of the
// picture.
type decodeNonVisualPictureProperties struct {
	CommonNonVisualProperties        *decodeCommonNonVisualProperties  `xml:"cNvPr"`
	CommonNonVisualPictureProperties *CommonNonVisualPictureProperties `xml:"cNvPicPr"`
	NonVisualProperties              *decodeNonVisualProperties        `xml:"nvPr"`
}

// decodeConnector directly maps the cxnSp element in the shape tree.
type decodeConnector struct {
	NonVisualConnectorProperties *decodeNonVisualConnectorProperties `xml:"nvCxnSpPr"`
	ShapeProperties              *DecodeShapeProperties              `xml:"spPr"`
}

// decodeNonVisualConnectorProperties directly maps the nvCxnSpPr element of
// the connector.
type decodeNonVisualConnectorProperties struct {
	CommonNonVisualProperties          *decodeCommonNonVisualProperties    `xml:"cNvPr"`
	CommonNonVisualConnectorProperties *CommonNonVisualConnectorProperties `xml:"cNvCxnSpPr"`
	NonVisualProperties                *decodeNonVisualProperties          `xml:"nvPr"`
}

// decodeCommonNonVisualProperties directly maps the cNvPr element of the
// non-visual properties.
type decodeCommonNonVisualProperties struct {
	ID             int              `xml:"id,attr"`
	Name           string           `xml:"name,attr"`
	Title          string           `xml:"title,attr,omitempty"`
	HlinkClick     *DecodeHyperlink `xml:"hlinkClick,omitempty"`
	HlinkMouseOver *DecodeHyperlink `xml:"hlinkMouseOver,omitempty"`
	ExtLst         *innerXML        `xml:"extLst,omitempty"`
}

// decodeBlipFill directly maps the blipFill element of the picture.
type decodeBlipFill struct {
	Blip    *decodeBlip    `xml:"blip"`
	Stretch *decodeStretch `xml:"stretch"`
}

// decodeBlip directly maps the blip element of the blip fill.
type decodeBlip struct {
	Embed       string       `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships embed,attr"`
	AlphaModFix *AlphaModFix `xml:"alphaModFix,omitempty"`
}

// decodeStretch directly maps the stretch element of the blip fill.
type decodeStretch struct {
	FillRect *FillRect `xml:"fillRect"`
}

type decodeNonVisualShapeProperties struct {
	CommonNonVisualProperties      *decodeCommonNonVisualProperties      `xml:"cNvPr"`
	CommonNonVisualShapeProperties *decodeCommonNonVisualShapeProperties `xml:"cNvSpPr"`
	NonVisualProperties            *decodeNonVisualProperties            `xml:"nvPr"`
}

type decodeCommonNonVisualShapeProperties struct {
	TxBox      *bool       `xml:"txBox,attr,omitempty"`
	ShapeLocks *ShapeLocks `xml:"spLocks,omitempty"`
}

type ShapeLocks struct {
	NoGroup *int `xml:"noGrp,attr,omitempty"`
}

type DecodeShapeProperties struct {
	Xfrm           *DecodeXfrm           `xml:"xfrm"`
	CustGeom       *innerXML             `xml:"custGeom,omitempty"`
	PresetGeometry *DecodePresetGeometry `xml:"prstGeom,omitempty"`
	NoFill         *noFill               `xml:"noFill,omitempty"`
	SolidFill      *DecodeSolidFill      `xml:"solidFill,omitempty"`
	Ln             *DecodeLine           `xml:"ln,omitempty"`
	EffectList     *innerXML             `xml:"effectLst,omitempty"`
}

type noFill struct{}

type DecodePresetGeometry struct {
	Preset          string           `xml:"prst,attr"`
	AdjustValueList *AdjustValueList `xml:"avLst"`
}

type AdjustValueList struct {
}

type AdjustValue struct {
	Name    string `xml:"name,attr"`
	Formula string `xml:"fmla,attr"`
}

// DecodeLine directly maps the ln element of the shape properties.
type DecodeLine struct {
	Width     *int             `xml:"w,attr,omitempty"`
	NoFill    *noFill          `xml:"noFill,omitempty"`
	SolidFill *DecodeSolidFill `xml:"solidFill,omitempty"`
	PrstDash  *PrstDash        `xml:"prstDash,omitempty"`
}

type DecodeTextBody struct {
	BodyProperties *DecodeBodyProperties `xml:"bodyPr"`
	ListStyle      *struct{}             `xml:"lstStyle,omitempty"`
	Paragraph      []DecodeParagraph     `xml:"p"`
}

type DecodeBodyProperties struct {
	AnchorCtr        *int         `xml:"anchorCtr,attr,omitempty"`
	Anchor           *string      `xml:"anchor,attr,omitempty"`
	BIns             *int         `xml:"bIns,attr,omitempty"`
	LIns             *int         `xml:"lIns,attr,omitempty"`
	SpcFirstLastPara *int         `xml:"spcFirstLastPara,attr,omitempty"`
	RIns             *int         `xml:"rIns,attr,omitempty"`
	Wrap             *string      `xml:"wrap,attr,omitempty"`
	TIns             *int         `xml:"tIns,attr,omitempty"`
	NormAutofit      *NormAutofit `xml:"normAutofit,omitempty"`
	NoAutofit        *NoAutofit   `xml:"noAutofit,omitempty"`
}

type NoAutofit struct{}

type DecodeParagraph struct {
	ParagraphProperties       *DecodeParagraphProperties `xml:"pPr,omitempty"`
	Runs                      []DecodeRuns               `xml:"r"`
	EndParagraphRunProperties *DecodeRunProperties       `xml:"endParaRPr,omitempty"`
}

// DecodeParagraphProperties directly maps the pPr element of the paragraph.
type DecodeParagraphProperties struct {
	Indent           *int                 `xml:"indent,attr,omitempty"`
	Lvl              *int                 `xml:"lvl,attr,omitempty"`
	MarL             *int                 `xml:"marL,attr,omitempty"`
	RTL              *int                 `xml:"rtl,attr,omitempty"`
	Align            *string              `xml:"algn,attr,omitempty"`
	LineSpacing      *DecodeLineSpacing   `xml:"lnSpc,omitempty"`
	SpacingBefore    *DecodeSpacingChange `xml:"spcBef,omitempty"`
	SpacingAfter     *DecodeSpacingChange `xml:"spcAft,omitempty"`
	BuFont           *DecodeBulletFont    `xml:"buFont,omitempty"`
	BuChar           *DecodeBulletChar    `xml:"buChar,omitempty"`
	BuNone           *struct{}            `xml:"buNone,omitempty"`
	DefRunProperties *DecodeRunProperties `xml:"defRPr,omitempty"`
}

// DecodeBulletFont directly maps the buFont element of the paragraph
// properties: the typeface of the bullet character.
type DecodeBulletFont struct {
	Typeface string `xml:"typeface,attr,omitempty"`
}

// DecodeBulletChar directly maps the buChar element of the paragraph
// properties: the character used as the bullet.
type DecodeBulletChar struct {
	Char string `xml:"char,attr,omitempty"`
}

// DecodeLineSpacing directly maps the lnSpc element of the paragraph
// properties.
type DecodeLineSpacing struct {
	SpacingPercent *SpacingPercent `xml:"spcPct,omitempty"`
}

// DecodeSpacingChange directly maps the spcBef or spcAft element of the
// paragraph properties.
type DecodeSpacingChange struct {
	SpacingPercent *SpacingPercent `xml:"spcPct,omitempty"`
	SpacingPts     *SpacingPts     `xml:"spcPts,omitempty"`
}

// SpacingPts defines the spacing in points (a:spcPts) element.
type SpacingPts struct {
	Val int `xml:"val,attr"`
}

type LineSpacing struct {
	SpacingPercent *SpacingPercent `xml:"a:spcPct"`
}

type SpacingPercent struct {
	Val int `xml:"val,attr"`
}

type DecodeRuns struct {
	RunProperties *DecodeRunProperties `xml:"rPr,omitempty"`
	Text          string               `xml:"t"`
}

type DecodeRunProperties struct {
	Bold           *int             `xml:"b,attr,omitempty"`
	Lang           string           `xml:"lang,attr,omitempty"`
	Size           *int             `xml:"sz,attr,omitempty"`
	Space          *int             `xml:"spc,attr,omitempty"`
	Strike         string           `xml:"strike,attr,omitempty"`
	SolidFill      *DecodeSolidFill `xml:"solidFill,omitempty"`
	Latin          *Font            `xml:"latin,omitempty"`
	EastAsian      *Font            `xml:"ea,omitempty"`
	ComplexScript  *Font            `xml:"cs,omitempty"`
	HlinkClick     *DecodeHyperlink `xml:"hlinkClick,omitempty"`
	HlinkMouseOver *DecodeHyperlink `xml:"hlinkMouseOver,omitempty"`
}

type DecodeSolidFill struct {
	SolidRGBColor *SolidRGBColor `xml:"srgbClr"`
}

type SolidRGBColor struct {
	Val string `xml:"val,attr"`
}

// Hyperlink defines the hyperlink (a:hlinkClick or a:hlinkMouseOver) element
// of the run properties or of the common non-visual properties: the id of the
// hyperlink relationship in the slide rels part and the tooltip shown on
// hover.
type Hyperlink struct {
	ID      string `xml:"r:id,attr"`
	Tooltip string `xml:"tooltip,attr,omitempty"`
}

// DecodeHyperlink directly maps the hlinkClick or hlinkMouseOver element of
// the run properties or of the common non-visual properties.
type DecodeHyperlink struct {
	ID      string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
	Tooltip string `xml:"tooltip,attr,omitempty"`
}

// Font defines a font typeface element (a:latin, a:ea or a:cs) of the run
// properties.
type Font struct {
	Typeface string `xml:"typeface,attr"`
}
