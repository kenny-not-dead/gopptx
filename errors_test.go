// Copyright 2026 kenny-not-dead. All rights reserved.
// Use of this source code is governed by a BSD-style license
// that can be found in the LICENSE file.
//
// Package gopptx provides functionality to create and manipulate PowerPoint
// (.pptx) files in Go, using the Office Open XML (ECMA-376) format.

package gopptx

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

// TestErrorMessages checks the messages of the errors of the package.
func TestErrorMessages(t *testing.T) {
	tests := []struct {
		name string
		err  error
		msg  string
	}{
		{"password length invalid", ErrPasswordLengthInvalid, "password length invalid"},
		{"unzip size limit", ErrOptionsUnzipSizeLimit,
			"the value of UnzipSizeLimit should be greater than or equal to UnzipXMLSizeLimit"},
		{"max file path length", ErrMaxFilePathLength, ErrMaxFilePathLength.Error()},
		{"save", ErrSave, "no path defined for file, consider File.WriteTo or File.Write"},
		{"slide name blank", ErrSlideNameBlank, "the slide name can not be blank"},
		{"presentation file format", ErrPresentationFileFormat, "unsupported presentation file format"},
		{"image empty", ErrImageEmpty, "the image data is empty"},
		{"unsupported image format", ErrUnsupportedImageFormat, "unsupported image format"},
		{"shape order invalid", ErrShapeOrderInvalid,
			"the order of the shapes does not match the shape tree of the slide"},
		{"invalid link type", ErrInvalidLinkType,
			"the link type must be LinkTypeExternal, LinkTypeLocation or LinkTypeNone"},
		{"shape hyperlink ambiguous", ErrShapeHyperlinkAmbiguous,
			"the shape contains runs with different hyperlinks"},
		{"table invalid size", ErrTableInvalidSize,
			"the table must have at least one row and one column"},
		{"chart empty", ErrChartEmpty, "the chart data is empty"},
		{"chart invalid data", ErrChartInvalidData,
			"the chart must have at least one category and one series with the value of each category"},
		{"slide not exist", ErrSlideNotExist{SlideID: 257}, "slide 257 does not exist"},
		{"slide layout not exist", ErrSlideLayoutNotExist{LayoutID: 104},
			"slide layout 104 does not exist"},
		{"slide master not exist", ErrSlideMasterNotExist{MasterID: 2147483648},
			"slide master 2147483648 does not exist"},
		{"slide master unsupported", ErrSlideMasterUnsupported{Master: "ppt/slideMasters/slideMaster1.xml",
			Element: "the shape of the type other than sp"},
			`the slide master ppt/slideMasters/slideMaster1.xml contains the unsupported element "the shape of the type other than sp"`},
		{"shape not exist", ErrShapeNotExist{ShapeID: 7}, "shape 7 does not exist"},
		{"picture not exist", ErrPictureNotExist{PictureID: 9}, "picture 9 does not exist"},
		{"connector not exist", ErrConnectorNotExist{ConnectorID: 11}, "connector 11 does not exist"},
		{"table not exist", ErrTableNotExist{TableID: 13}, "table 13 does not exist"},
		{"chart not exist", ErrChartNotExist{ChartID: 15}, "chart 15 does not exist"},
		{"table cell invalid", ErrTableCellInvalid{Row: 2, Col: 3}, "the table cell (2, 3) does not exist"},
		{"notes slide not exist", ErrNotesSlideNotExist{SlideID: 256}, "the slide 256 has no notes"},
		{"notes slide exists", ErrNotesSlideExists{SlideID: 256}, "the slide 256 already has notes"},
		{"notes body placeholder not exist", ErrNotesBodyPlaceholderNotExist,
			"the notes slide has no body placeholder"},
		{"image not exist", ErrImageNotExist{RelationshipID: "rId2"},
			"image relationship rId2 does not exist"},
		{"link target invalid", ErrLinkTargetInvalid{Target: "abc"},
			`the link target "abc" is not a slide id`},
		{"slide layout unsupported", ErrSlideLayoutUnsupported{Layout: "ppt/slideLayouts/slideLayout1.xml",
			Element: "the timing element of the slide layout"},
			`the slide layout ppt/slideLayouts/slideLayout1.xml contains the unsupported element "the timing element of the slide layout"`},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			assert.EqualError(t, tt.err, tt.msg)
		})
	}

	assert.EqualError(t, newUnzipSizeLimitError(10485760), "unzip size exceeds the 10485760 bytes limit")
	assert.EqualError(t, unexpectedNamespace("http://schemas.example.com"),
		"Unexpected namespace: http://schemas.example.com")
}
