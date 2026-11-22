// Copyright 2026 kenny-not-dead. All rights reserved.
// Use of this source code is governed by a BSD-style license
// that can be found in the LICENSE file.
//
// Package gopptx provides functionality to create and manipulate PowerPoint
// (.pptx) files in Go, using the Office Open XML (ECMA-376) format.

package gopptx

import (
	"io"
	"os"
	"path/filepath"
	"testing"

	"github.com/stretchr/testify/assert"
)

// TestBufferedWriter checks the bufferedWriter: the buffer without the temp
// file (the reads and the writes are in memory), the buffer with the temp
// file (the flush writes the buffer to the temp file, the reader reads the
// temp file) and the close of the buffer (the temp file is removed).
func TestBufferedWriter(t *testing.T) {
	// the buffer without the temp file
	var bw bufferedWriter
	_, err := bw.buf.Write([]byte("in memory"))
	assert.NoError(t, err)

	reader, err := bw.Reader()
	assert.NoError(t, err)
	data, err := io.ReadAll(reader)
	assert.NoError(t, err)
	assert.Equal(t, []byte("in memory"), data)

	assert.NoError(t, bw.Flush())
	assert.NoError(t, bw.Close())
	assert.Empty(t, bw.buf.String())

	// the buffer with the temp file
	tmpDir := t.TempDir()
	tmp, err := os.CreateTemp(tmpDir, "buffered")
	assert.NoError(t, err)
	bw = bufferedWriter{tmpDir: tmpDir, tmp: tmp}
	_, err = bw.buf.Write([]byte("to the file"))
	assert.NoError(t, err)

	assert.NoError(t, bw.Flush())
	assert.Empty(t, bw.buf.String())

	reader, err = bw.Reader()
	assert.NoError(t, err)
	data, err = io.ReadAll(reader)
	assert.NoError(t, err)
	assert.Equal(t, []byte("to the file"), data)

	assert.NoError(t, bw.Close())
	_, err = os.Stat(tmp.Name())
	assert.True(t, os.IsNotExist(err), "the temp file is not removed by the close")

	// the reader of the closed buffer returns the error
	_, err = bw.Reader()
	assert.Error(t, err)

	// the flush and the close of the buffer without the temp file
	var empty bufferedWriter
	assert.NoError(t, empty.Flush())
	assert.NoError(t, empty.Close())
}

// TestStreamWriterFile checks that the presentation with the slides that are
// read and saved keeps the content (the stream of the slide part is closed
// and removed on the close of the file).
func TestStreamWriterFile(t *testing.T) {
	f, err := OpenFile(testFixture)
	assert.NoError(t, err)

	_, err = f.GetShapes(257)
	assert.NoError(t, err)

	out := filepath.Join(t.TempDir(), "stream.pptx")
	assert.NoError(t, f.SaveAs(out))
	assert.NoError(t, assertPartsEqual(t, testFixture, out))
	assert.NoError(t, f.Close())
}
