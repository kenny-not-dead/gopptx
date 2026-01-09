// Copyright 2026 kenny-not-dead. All rights reserved.
// Use of this source code is governed by a BSD-style license
// that can be found in the LICENSE file.
//
// Package gopptx provides functionality to create and manipulate PowerPoint
// (.pptx) files in Go, using the Office Open XML (ECMA-376) format.

package gopptx

import (
	_ "embed"
	"encoding/xml"
)

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
	MaxFilePathLength = 207
	StreamChunkSize   = 1 << 24
	UnzipSizeLimit    = 1000 << 24
)

const (
	ContentTypePresentationML                     = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
	ContentTypeSlideML                            = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
	ContentTypeVBA                                = "application/vnd.ms-office.vbaProject"
	NameSpaceDrawingMLMain                        = "http://schemas.openxmlformats.org/drawingml/2006/main"
	NameSpaceExtendedProperties                   = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
	NameSpaceXML                                  = "http://www.w3.org/XML/1998/namespace"
	SourceRelationshipCustomProperties            = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
	SourceRelationshipOfficeDocument              = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
	StrictNameSpaceDocumentPropertiesVariantTypes = "http://purl.oclc.org/ooxml/officeDocument/docPropsVTypes"
	StrictNameSpaceDrawingMLMain                  = "http://purl.oclc.org/ooxml/drawingml/main"
	StrictNameSpaceExtendedProperties             = "http://purl.oclc.org/ooxml/officeDocument/extendedProperties"

	NameSpacePresentationMLMain       = "http://schemas.openxmlformats.org/presentationml/2006/main"
	StrictNameSpacePresentationMLMain = "http://purl.oclc.org/ooxml/presentationml/main"
)

const (
	defaultXMLPathContentTypes     = "[Content_Types].xml"
	defaultXMLPathDocPropsApp      = "docProps/app.xml"
	defaultXMLPathDocPropsCore     = "docProps/core.xml"
	defaultXMLPathPresentation     = "ppt/presentation.xml"
	defaultXMLPathPresProps        = "ppt/presProps.xml"
	defaultXMLPathPresentationRels = "ppt/_rels/presentation.xml.rels"
	defaultXMLPathRels             = "_rels/.rels"
	defaultXMLPathSlide            = "ppt/slides/slide1.xml"
	defaultXMLPathSlideRels        = "ppt/slides/_rels/slide1.xml.rels"
	defaultXMLPathSlideLayout      = "ppt/slideLayouts/slideLayout1.xml"
	defaultXMLPathSlideLayoutRels  = "ppt/slideLayouts/_rels/slideLayout1.xml.rels"
	defaultXMLPathSlideMaster      = "ppt/slideMasters/slideMaster1.xml"
	defaultXMLPathSlideMasterRels  = "ppt/slideMasters/_rels/slideMaster1.xml.rels"
	defaultXMLPathTheme            = "ppt/theme/theme1.xml"
)

const (
	defaultXMLThemeID       = "rId1"
	defaultXMLMasterSlideID = "rId2"
	defaultXMLSlideID       = "rId3"
)

const (
	MaxFieldLength = 255
)

var (
	//go:embed templates/\[Content_Types].xml
	templateContentTypes string

	//go:embed templates/\app.xml
	templateDocpropsApp string

	//go:embed templates/core.xml
	templateDocpropsCore string

	//go:embed templates/slide1.xml
	TemplateSlide string

	//go:embed templates/slide1-rels.xml
	TemplateSlideRels string

	//go:embed templates/slideLayout1.xml
	templateSlideLayout string

	//go:embed templates/slideLayout1-rels.xml
	templateSlideLayoutRels string

	//go:embed templates/slideMaster1.xml
	templateSlideMaster string

	//go:embed templates/slideMaster1-rels.xml
	templateSlideMasterRels string

	//go:embed templates/presentation.xml
	templatePresentation string

	//go:embed templates/presProps.xml
	templatePresProps string

	//go:embed templates/presentation-rels.xml
	templatePresentationRels string

	//go:embed templates/rels.xml
	templateRels string

	//go:embed templates/theme1.xml
	templateTheme string
)

// supportedContentTypes defined supported file format types.
var supportedContentTypes = map[string]string{
	".pptx": ContentTypePresentationML,
}

const (
	templateNamespaceIDMap = ` xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main"`
	templatePPTXNamespace  = ` xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="p14" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"`
)
