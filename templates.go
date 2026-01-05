// Copyright 2026 kenny-not-dead. All rights reserved.
// Use of this source code is governed by a BSD-style license
// that can be found in the LICENSE file.
//
// Package gopptx provides functionality to create and manipulate PowerPoint
// (.pptx) files in Go, using the Office Open XML (ECMA-376) format.

package gopptx

import "encoding/xml"

var (
	NameSpaceDocumentPropertiesVariantTypes = xml.Attr{Name: xml.Name{Local: "vt", Space: "xmlns"}, Value: "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"}
	NameSpaceDrawing2016SVG                 = xml.Attr{Name: xml.Name{Local: "asvg", Space: "xmlns"}, Value: "http://schemas.microsoft.com/office/drawing/2016/SVG/main"}
	NameSpaceDrawingML                      = xml.Attr{Name: xml.Name{Local: "a", Space: "xmlns"}, Value: "http://schemas.openxmlformats.org/drawingml/2006/main"}
	NameSpaceDrawingMLA14                   = xml.Attr{Name: xml.Name{Local: "a14", Space: "xmlns"}, Value: "http://schemas.microsoft.com/office/drawing/2010/main"}
	NameSpaceDrawingMLChart                 = xml.Attr{Name: xml.Name{Local: "c", Space: "xmlns"}, Value: "http://schemas.openxmlformats.org/drawingml/2006/chart"}
	NameSpaceDrawingMLSlicer                = xml.Attr{Name: xml.Name{Local: "sle", Space: "xmlns"}, Value: "http://schemas.microsoft.com/office/drawing/2010/slicer"}
	NameSpaceDrawingMLSlicerX15             = xml.Attr{Name: xml.Name{Local: "sle15", Space: "xmlns"}, Value: "http://schemas.microsoft.com/office/drawing/2012/slicer"}
	SourceRelationship                      = xml.Attr{Name: xml.Name{Local: "r", Space: "xmlns"}, Value: "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
	SourceRelationshipCompatibility         = xml.Attr{Name: xml.Name{Local: "mc", Space: "xmlns"}, Value: "http://schemas.openxmlformats.org/markup-compatibility/2006"}
	NameSpacePresentationML                 = xml.Attr{Name: xml.Name{Local: "p", Space: "xmlns"}, Value: "http://schemas.openxmlformats.org/presentationml/2006/main"}
)

const (
	StreamChunkSize = 1 << 24
	UnzipSizeLimit  = 1000 << 24
)

const (
	NameSpaceDrawingMLMain                        = "http://schemas.openxmlformats.org/drawingml/2006/main"
	NameSpaceExtendedProperties                   = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
	SourceRelationshipCustomProperties            = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
	SourceRelationshipOfficeDocument              = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
	StrictNameSpaceDocumentPropertiesVariantTypes = "http://purl.oclc.org/ooxml/officeDocument/docPropsVTypes"
	StrictNameSpaceDrawingMLMain                  = "http://purl.oclc.org/ooxml/drawingml/main"
	StrictNameSpaceExtendedProperties             = "http://purl.oclc.org/ooxml/officeDocument/extendedProperties"

	NameSpacePresentationMLMain       = "http://schemas.openxmlformats.org/presentationml/2006/main"
	StrictNameSpacePresentationMLMain = "http://purl.oclc.org/ooxml/presentationml/main"
)

const (
	defaultXMLPathContentTypes = "[Content_Types].xml"
	defaultXMLPathRels         = "_rels/.rels"
)

const (
	MaxFieldLength = 255
)
