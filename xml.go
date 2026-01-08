// Copyright 2026 kenny-not-dead. All rights reserved.
// Use of this source code is governed by a BSD-style license
// that can be found in the LICENSE file.
//
// Package gopptx provides functionality to create and manipulate PowerPoint
// (.pptx) files in Go, using the Office Open XML (ECMA-376) format.

package gopptx

// innerXML holds parts of XML content currently not unmarshal.
type innerXML struct {
	Content string `xml:",innerxml"`
}

// alternateContent is a container for a sequence of multiple
// representations of a given piece of content. The program reading the file
// should only process one of these, and the one chosen should be based on
// which conditions match.
type alternateContent struct {
	XMLNSMC string `xml:"xmlns:mc,attr,omitempty"`
	Content string `xml:",innerxml"`
}
