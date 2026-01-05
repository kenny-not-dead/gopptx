// Copyright 2026 kenny-not-dead. All rights reserved.
// Use of this source code is governed by a BSD-style license
// that can be found in the LICENSE file.
//
// Package gopptx provides functionality to create and manipulate PowerPoint
// (.pptx) files in Go, using the Office Open XML (ECMA-376) format.

package gopptx

import "os"

// Close closes and cleanup the open temporary file for the spreadsheet.
func (f *File) Close() error {
	var firstErr error

	for _, stream := range f.streams {
		_ = stream.rawData.Close()
	}
	f.streams = nil
	f.tempFiles.Range(func(k, v interface{}) bool {
		if err := os.Remove(v.(string)); err != nil && firstErr == nil {
			firstErr = err
		}
		return true
	})
	f.tempFiles.Clear()
	return firstErr
}
