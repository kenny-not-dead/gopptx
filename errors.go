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
	// ErrUnsupportedEncryptMechanism defined the error message on receive the blank slide name.
	ErrSlideNameBlank = errors.New("the slide name can not be blank")
	// ErrPresentationFileFormat defined the error message on receive an
	// unsupported presentation file format.
	ErrPresentationFileFormat = errors.New("unsupported presentation file format")
)

// ErrSlideNotExist defined an error of slide that does not exist.
type ErrSlideNotExist struct {
	SlideID int
}

// Error returns the error message on receiving the non existing slide name.
func (err ErrSlideNotExist) Error() string {
	return fmt.Sprintf("slide %d does not exist", err.SlideID)
}

// newUnzipSizeLimitError defined the error message on unzip size exceeds the
// limit.
func newUnzipSizeLimitError(unzipSizeLimit int64) error {
	return fmt.Errorf("unzip size exceeds the %d bytes limit", unzipSizeLimit)
}

// TODO
func unexpectedNamespace(space string) error {
	return fmt.Errorf("Unexpected namespace: %s", space)
}
