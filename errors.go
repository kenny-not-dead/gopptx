// Copyright 2026 kenny-not-dead. All rights reserved.
// Use of this source code is governed by a BSD-style license
// that can be found in the LICENSE file.
//
// Package gopptx provides functionality to create and manipulate PowerPoint
// (.pptx) files in Go, using the Office Open XML (ECMA-376) format.

package gopptx

import (
	"errors"
	"fmt"
)

var (
	// ErrPasswordLengthInvalid defined the error message on invalid password
	// length.
	ErrPasswordLengthInvalid = errors.New("password length invalid")
	// ErrOptionsUnzipSizeLimit defined the error message for receiving
	// invalid UnzipSizeLimit and UnzipXMLSizeLimit.
	ErrOptionsUnzipSizeLimit = errors.New("the value of UnzipSizeLimit should be greater than or equal to UnzipXMLSizeLimit")
	// ErrMaxFilePathLength defined the error message on receive the file path
	// length overflow.
	ErrMaxFilePathLength = fmt.Errorf("file path length exceeds maximum limit %d characters", MaxFilePathLength)
	// ErrSave defined the error message for saving file.
	ErrSave = errors.New("no path defined for file, consider File.WriteTo or File.Write")
	// ErrSlideNameBlank is returned when the slide name is empty.
	ErrSlideNameBlank = errors.New("the slide name can not be blank")
	// ErrPresentationFileFormat defined the error message on receive an
	// unsupported presentation file format.
	ErrPresentationFileFormat = errors.New("unsupported presentation file format")
	// ErrImageEmpty defined the error message on receive the empty image data.
	ErrImageEmpty = errors.New("the image data is empty")
	// ErrUnsupportedImageFormat defined the error message on receive an
	// unsupported image format.
	ErrUnsupportedImageFormat = errors.New("unsupported image format")
	// ErrShapeOrderInvalid is returned when the given order of the shape
	// tree children does not match the shape tree of the slide: every
	// child of the shape tree must be contained in the order exactly once.
	ErrShapeOrderInvalid = errors.New("the order of the shapes does not match the shape tree of the slide")
	// ErrInvalidLinkType is returned when the given hyperlink link type is
	// not supported: the link type must be LinkTypeExternal or LinkTypeNone.
	ErrInvalidLinkType = errors.New("the link type must be LinkTypeExternal or LinkTypeNone")
	// ErrShapeHyperlinkAmbiguous is returned when the shape contains runs
	// with different hyperlinks, so the hyperlink of the shape is ambiguous.
	ErrShapeHyperlinkAmbiguous = errors.New("the shape contains runs with different hyperlinks")
	// ErrTableInvalidSize is returned when the given size of the new table
	// is invalid: the table must have at least one row and one column.
	ErrTableInvalidSize = errors.New("the table must have at least one row and one column")
)

// ErrSlideNotExist defined an error of slide that does not exist.
type ErrSlideNotExist struct {
	SlideID int
}

// Error returns the error message on receiving the non existing slide name.
func (err ErrSlideNotExist) Error() string {
	return fmt.Sprintf("slide %d does not exist", err.SlideID)
}

// newUnzipSizeLimitError returns an error when the unzipped content exceeds the specified limit.
func newUnzipSizeLimitError(unzipSizeLimit int64) error {
	return fmt.Errorf("unzip size exceeds the %d bytes limit", unzipSizeLimit)
}

// unexpectedNamespace returns an error when an unexpected XML namespace is encountered.
func unexpectedNamespace(space string) error {
	return fmt.Errorf("Unexpected namespace: %s", space)
}

// ErrShapeNotExist defined an error of shape that does not exist.
type ErrShapeNotExist struct {
	ShapeID int
}

// Error returns the error message on receiving the non existing shape.
func (err ErrShapeNotExist) Error() string {
	return fmt.Sprintf("shape %d does not exist", err.ShapeID)
}

// ErrPictureNotExist defined an error of picture that does not exist.
type ErrPictureNotExist struct {
	PictureID int
}

// Error returns the error message on receiving the non existing picture.
func (err ErrPictureNotExist) Error() string {
	return fmt.Sprintf("picture %d does not exist", err.PictureID)
}

// ErrConnectorNotExist defined an error of connector that does not exist.
type ErrConnectorNotExist struct {
	ConnectorID int
}

// Error returns the error message on receiving the non existing connector.
func (err ErrConnectorNotExist) Error() string {
	return fmt.Sprintf("connector %d does not exist", err.ConnectorID)
}

// ErrTableNotExist defined an error of table that does not exist.
type ErrTableNotExist struct {
	TableID int
}

// Error returns the error message on receiving the non existing table.
func (err ErrTableNotExist) Error() string {
	return fmt.Sprintf("table %d does not exist", err.TableID)
}

// ErrTableCellInvalid defined an error of a table cell that is out of the
// bounds of the table.
type ErrTableCellInvalid struct {
	Row int
	Col int
}

// Error returns the error message on receiving the table cell out of the
// bounds of the table.
func (err ErrTableCellInvalid) Error() string {
	return fmt.Sprintf("the table cell (%d, %d) does not exist", err.Row, err.Col)
}

// ErrNotesSlideNotExist defined an error of a slide that has no notes slide.
type ErrNotesSlideNotExist struct {
	SlideID int
}

// Error returns the error message on receiving the slide that has no notes
// slide.
func (err ErrNotesSlideNotExist) Error() string {
	return fmt.Sprintf("the slide %d has no notes", err.SlideID)
}

// ErrNotesSlideExists defined an error of a slide that already has a notes
// slide.
type ErrNotesSlideExists struct {
	SlideID int
}

// Error returns the error message on receiving the slide that already has a
// notes slide.
func (err ErrNotesSlideExists) Error() string {
	return fmt.Sprintf("the slide %d already has notes", err.SlideID)
}

// ErrNotesBodyPlaceholderNotExist defined an error of a notes slide that has
// no notes body placeholder.
var ErrNotesBodyPlaceholderNotExist = errors.New("the notes slide has no body placeholder")

// ErrImageNotExist defined an error of image relationship that does not exist.
type ErrImageNotExist struct {
	RelationshipID string
}

// Error returns the error message on receiving the non existing image
// relationship.
func (err ErrImageNotExist) Error() string {
	return fmt.Sprintf("image relationship %s does not exist", err.RelationshipID)
}
