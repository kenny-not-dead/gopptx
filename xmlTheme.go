// Copyright 2026 kenny-not-dead. All rights reserved.
// Use of this source code is governed by a BSD-style license
// that can be found in the LICENSE file.
//
// Package gopptx provides functionality to create and manipulate PowerPoint
// (.pptx) files in Go, using the Office Open XML (ECMA-376) format.

package gopptx

import "encoding/xml"

// theme directly maps the theme element in the namespace
// http://schemas.openxmlformats.org/drawingml/2006/main
type theme struct {
	XMLName       xml.Name   `xml:"a:theme"`
	XMLNSa        string     `xml:"xmlns:a,attr"`
	XMLNSr        string     `xml:"xmlns:r,attr"`
	Name          string     `xml:"name,attr"`
	ThemeElements baseStyles `xml:"a:themeElements"`
}

// baseStyles defines the theme elements for a theme, and is the workhorse
// of the theme. The bulk of the shared theme information that is used by a
// given document is defined here. Within this complex type is defined a color
// scheme, a font scheme, and a style matrix (format scheme) that defines
// different formatting options for different pieces of a document.
type baseStyles struct {
	ColorScheme  colorScheme `xml:"a:clrScheme"`
	FontScheme   fontScheme  `xml:"a:fontScheme"`
	FormatScheme styleMatrix `xml:"a:fmtScheme"`
}

// colorScheme defines a set of colors for the theme. The set of colors
// consists of twelve color slots that can each hold a color of choice.
type colorScheme struct {
	Name     string           `xml:"name,attr"`
	Dk1      complexTypeColor `xml:"a:dk1"`
	Lt1      complexTypeColor `xml:"a:lt1"`
	Dk2      complexTypeColor `xml:"a:dk2"`
	Lt2      complexTypeColor `xml:"a:lt2"`
	Accent1  complexTypeColor `xml:"a:accent1"`
	Accent2  complexTypeColor `xml:"a:accent2"`
	Accent3  complexTypeColor `xml:"a:accent3"`
	Accent4  complexTypeColor `xml:"a:accent4"`
	Accent5  complexTypeColor `xml:"a:accent5"`
	Accent6  complexTypeColor `xml:"a:accent6"`
	Hlink    complexTypeColor `xml:"a:hlink"`
	FolHlink complexTypeColor `xml:"a:folHlink"`
}

// complexTypeColor holds the actual color values that are to be applied to a given
// diagram and how those colors are to be applied.
type complexTypeColor struct {
	ScrgbClr    *innerXML    `xml:"a:scrgbClr"`
	SrgbColor   *srgbColor   `xml:"a:srgbClr"`
	HslClr      *innerXML    `xml:"a:hslClr"`
	SystemColor *systemColor `xml:"a:sysClr"`
	SchemeColor *innerXML    `xml:"a:schemeClr"`
	PresetColor *innerXML    `xml:"a:prstClr"`
}

// complexTypeSupplementalFont defines an additional font that is used for language
// specific fonts in themes. For example, one can specify a font that gets used
// only within the Japanese language context.
type complexTypeSupplementalFont struct {
	Script   string `xml:"script,attr"`
	Typeface string `xml:"typeface,attr"`
}

// fontCollection defines a major and minor font which is used in the font
// scheme. A font collection consists of a font definition for Latin, East
// Asian, and complex script. On top of these three definitions, one can also
// define a font for use in a specific language or languages.
type fontCollection struct {
	Latin *complexTypeTextFont          `xml:"a:latin"`
	Ea    *complexTypeTextFont          `xml:"a:ea"`
	Cs    *complexTypeTextFont          `xml:"a:cs"`
	Font  []complexTypeSupplementalFont `xml:"a:font"`
}

// fontScheme element defines the font scheme within the theme. The font
// scheme consists of a pair of major and minor fonts for which to use in a
// document. The major font corresponds well with the heading areas of a
// document, and the minor font corresponds well with the normal text or
// paragraph areas.
type fontScheme struct {
	Name      string         `xml:"name,attr"`
	MajorFont fontCollection `xml:"a:majorFont"`
	MinorFont fontCollection `xml:"a:minorFont"`
}

// styleMatrix defines a set of formatting options, which can be referenced
// by documents that apply a certain style to a given part of an object. For
// example, in a given shape, say a rectangle, one can reference a themed line
// style, themed effect, and themed fill that would be theme specific and
// change when the theme is changed.
type styleMatrix struct {
	Name            string          `xml:"name,attr,omitempty"`
	FillStyleList   fillStyleList   `xml:"a:fillStyleLst"`
	LineStyleList   lineStyleList   `xml:"a:lnStyleLst"`
	EffectStyleList effectStyleList `xml:"a:effectStyleLst"`
	BgFillStyleList bgFillStyleList `xml:"a:bgFillStyleLst"`
}

// fillStyleList element defines a set of three fill styles that are used
// within a theme. The three fill styles are arranged in order from subtle to
// moderate to intense.
type fillStyleList struct {
	FillStyleLst string `xml:",innerxml"`
}

// lineStyleList element defines a list of three line styles for use within a
// theme. The three line styles are arranged in order from subtle to moderate
// to intense versions of lines. This list makes up part of the style matrix.
type lineStyleList struct {
	LineStyleList string `xml:",innerxml"`
}

// effectStyleList element defines a set of three effect styles that create
// the effect style list for a theme. The effect styles are arranged in order
// of subtle to moderate to intense.
type effectStyleList struct {
	EffectStyleLst string `xml:",innerxml"`
}

// bgFillStyleList element defines a list of background fills that are
// used within a theme. The background fills consist of three fills, arranged
// in order from subtle to moderate to intense.
type bgFillStyleList struct {
	BgFillStyleLst string `xml:",innerxml"`
}

// systemColor element specifies a color bound to predefined operating system
// elements.
type systemColor struct {
	Val     string `xml:"val,attr"`
	LastClr string `xml:"lastClr,attr"`
}

// decodeTheme defines the structure used to parse the a:theme element for the
// theme.
type decodeTheme struct {
	XMLName       xml.Name         `xml:"http://schemas.openxmlformats.org/drawingml/2006/main theme"`
	Name          string           `xml:"name,attr"`
	ThemeElements decodeBaseStyles `xml:"themeElements"`
}

// decodeBaseStyles defines the structure used to parse the theme elements for a
// theme, and is the workhorse of the theme.
type decodeBaseStyles struct {
	ColorScheme  decodeColorScheme `xml:"clrScheme"`
	FontScheme   decodeFontScheme  `xml:"fontScheme"`
	FormatScheme decodeStyleMatrix `xml:"fmtScheme"`
}

// decodeColorScheme defines the structure used to parse a set of colors for the
// theme.
type decodeColorScheme struct {
	Name     string                      `xml:"name,attr"`
	Dk1      decodeComplexTypeColorColor `xml:"dk1"`
	Lt1      decodeComplexTypeColorColor `xml:"lt1"`
	Dk2      decodeComplexTypeColorColor `xml:"dk2"`
	Lt2      decodeComplexTypeColorColor `xml:"lt2"`
	Accent1  decodeComplexTypeColorColor `xml:"accent1"`
	Accent2  decodeComplexTypeColorColor `xml:"accent2"`
	Accent3  decodeComplexTypeColorColor `xml:"accent3"`
	Accent4  decodeComplexTypeColorColor `xml:"accent4"`
	Accent5  decodeComplexTypeColorColor `xml:"accent5"`
	Accent6  decodeComplexTypeColorColor `xml:"accent6"`
	Hlink    decodeComplexTypeColorColor `xml:"hlink"`
	FolHlink decodeComplexTypeColorColor `xml:"folHlink"`
}

// decodeFontScheme defines the structure used to parse font scheme within the
// theme.
type decodeFontScheme struct {
	Name      string               `xml:"name,attr"`
	MajorFont decodeFontCollection `xml:"majorFont"`
	MinorFont decodeFontCollection `xml:"minorFont"`
}

// decodeFontCollection defines the structure used to parse a major and minor
// font which is used in the font scheme.
type decodeFontCollection struct {
	Latin *complexTypeTextFont          `xml:"latin"`
	Ea    *complexTypeTextFont          `xml:"ea"`
	Cs    *complexTypeTextFont          `xml:"cs"`
	Font  []complexTypeSupplementalFont `xml:"font"`
}

// decodeComplexTypeColorColor defines the structure used to parse the actual color values
// that are to be applied to a given diagram and how those colors are to be
// applied.
type decodeComplexTypeColorColor struct {
	ScrgbColor  *innerXML    `xml:"scrgbClr"`
	SrgbColor   *srgbColor   `xml:"srgbClr"`
	HslColor    *innerXML    `xml:"hslClr"`
	SystemColor *systemColor `xml:"sysClr"`
	SchemeColor *innerXML    `xml:"schemeClr"`
	PresetColor *innerXML    `xml:"prstClr"`
}

// decodeStyleMatrix defines the structure used to parse a set of formatting
// options, which can be referenced by documents that apply a certain style to
// a given part of an object.
type decodeStyleMatrix struct {
	Name            string          `xml:"name,attr,omitempty"`
	FillStyleList   fillStyleList   `xml:"fillStyleLst"`
	LineStyleList   lineStyleList   `xml:"lnStyleLst"`
	EffectStyleList effectStyleList `xml:"effectStyleLst"`
	BgFillStyleList bgFillStyleList `xml:"bgFillStyleLst"`
}

// srgbColor directly maps the val element with string data type as an
// attribute.
type srgbColor struct {
	Val *string `xml:"val,attr"`
}

// innerXML holds parts of XML content currently not unmarshal.
type innerXML struct {
	Content string `xml:",innerxml"`
}

type complexTypeTextFont struct {
	Typeface    string `xml:"typeface,attr"`
	Panose      string `xml:"panose,attr,omitempty"`
	PitchFamily string `xml:"pitchFamily,attr,omitempty"`
	Charset     string `xml:"Charset,attr,omitempty"`
}
