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

// contentTypes directly maps the types' element of content types for relationship
// parts, it takes a Multipurpose Internet Mail Extension (MIME) media type as a
// value.
type contentTypes struct {
	mu        sync.Mutex
	XMLName   xml.Name              `xml:"http://schemas.openxmlformats.org/package/2006/content-types Types"`
	Defaults  []contentTypeDefault  `xml:"Default"`
	Overrides []contentTypeOverride `xml:"Override"`
}

// contentTypeOverride directly maps the override element in the namespace
// http://schemas.openxmlformats.org/package/2006/content-types
type contentTypeOverride struct {
	PartName    string `xml:",attr"`
	ContentType string `xml:",attr"`
}

// contentTypeDefault directly maps the default element in the namespace
// http://schemas.openxmlformats.org/package/2006/content-types
type contentTypeDefault struct {
	Extension   string `xml:",attr"`
	ContentType string `xml:",attr"`
}
