// Copyright 2026 kenny-not-dead. All rights reserved.
// Use of this source code is governed by a BSD-style license
// that can be found in the LICENSE file.
//
// Package gopptx provides functionality to create and manipulate PowerPoint
// (.pptx) files in Go, using the Office Open XML (ECMA-376) format.

package gopptx

import (
	"bytes"
	"io"
	"os"
)

type StreamWriter struct {
	file         *File
	SlideName    string
	SlideID      int
	slideWritten bool
	slide        *Slide
	rawData      bufferedWriter
}

// bufferedWriter uses a temp file to store an extended buffer. Writes are
// always made to an in-memory buffer, which will always succeed. The buffer
// is written to the temp file with Sync, which may return an error.
// Therefore, Sync should be periodically called and the error checked.
type bufferedWriter struct {
	tmpDir string
	tmp    *os.File
	buf    bytes.Buffer
}

// Close the underlying temp file and reset the in-memory buffer.
func (bw *bufferedWriter) Close() error {
	bw.buf.Reset()
	if bw.tmp == nil {
		return nil
	}
	defer os.Remove(bw.tmp.Name())
	return bw.tmp.Close()
}

// Reader provides read-access to the underlying buffer/file.
func (bw *bufferedWriter) Reader() (io.Reader, error) {
	if bw.tmp == nil {
		return bytes.NewReader(bw.buf.Bytes()), nil
	}
	if err := bw.Flush(); err != nil {
		return nil, err
	}
	fi, err := bw.tmp.Stat()
	if err != nil {
		return nil, err
	}
	// os.File.ReadAt does not affect the cursor position and is safe to use here
	return io.NewSectionReader(bw.tmp, 0, fi.Size()), nil
}

// Flush the entire in-memory buffer to the temp file, if a temp file is being
// used.
func (bw *bufferedWriter) Flush() error {
	if bw.tmp == nil {
		return nil
	}
	_, err := bw.buf.WriteTo(bw.tmp)
	if err != nil {
		return err
	}
	bw.buf.Reset()
	return nil
}
