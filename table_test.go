// Copyright 2026 kenny-not-dead. All rights reserved.
// Use of this source code is governed by a BSD-style license
// that can be found in the LICENSE file.
//
// Package gopptx provides functionality to create and manipulate PowerPoint
// (.pptx) files in Go, using the Office Open XML (ECMA-376) format.

package gopptx

import (
	"errors"
	"path/filepath"
	"strconv"
	"strings"
	"testing"
)

// cellText returns the plain text of the cell of the table.
func cellText(t *testing.T, table Table, row, col int) string {
	t.Helper()
	data := table.TableData
	if row < 0 || row >= len(data.TableRows) {
		t.Fatalf("row %d out of range", row)
	}
	cell := data.TableRows[row].TableCells[col]
	if cell.TextBody == nil {
		t.Fatalf("the cell (%d, %d) has no text body", row, col)
	}
	var sb strings.Builder
	for i := range cell.TextBody.Paragraph {
		for j := range cell.TextBody.Paragraph[i].Runs {
			sb.WriteString(cell.TextBody.Paragraph[i].Runs[j].Text)
		}
		if i != len(cell.TextBody.Paragraph)-1 {
			sb.WriteString("\n")
		}
	}
	return sb.String()
}

// TestAddTable checks that AddTable creates the table in the graphic frame
// (p:graphicFrame) and that the table can be read back.
func TestAddTable(t *testing.T) {
	f := NewFile()
	defer f.Close()
	slideID := f.GetSlideList()[0]

	// the size of the table must have at least one row and one column
	if _, err := f.AddTable(slideID, &TableOptions{}); !errors.Is(err, ErrTableInvalidSize) {
		t.Fatalf("expected ErrTableInvalidSize, got %v", err)
	}
	if _, err := f.AddTable(999, &TableOptions{Rows: 1, Cols: 1}); !errors.As(err, new(ErrSlideNotExist)) {
		t.Fatalf("expected ErrSlideNotExist, got %v", err)
	}

	tableID, err := f.AddTable(slideID, &TableOptions{Rows: 2, Cols: 3})
	if err != nil {
		t.Fatal(err)
	}

	tables, err := f.GetTables(slideID)
	if err != nil {
		t.Fatal(err)
	}
	if len(tables) != 1 {
		t.Fatalf("expected 1 table, got %d", len(tables))
	}
	if tables[0].ID != tableID {
		t.Fatalf("table id = %d, want %d", tables[0].ID, tableID)
	}
	if tables[0].Name != "Table "+strconv.Itoa(tableID) {
		t.Fatalf("table name = %q, want the default name", tables[0].Name)
	}

	table, err := f.GetTable(slideID, tableID)
	if err != nil {
		t.Fatal(err)
	}
	data := table.TableData
	if len(data.TableRows) != 2 {
		t.Fatalf("expected 2 rows, got %d", len(data.TableRows))
	}
	if len(data.TableGrid.GridCols) != 3 {
		t.Fatalf("expected 3 grid columns, got %d", len(data.TableGrid.GridCols))
	}
	if len(data.TableRows[0].TableCells) != 3 {
		t.Fatalf("expected 3 cells in the row, got %d", len(data.TableRows[0].TableCells))
	}
	if data.TableProperties == nil || data.TableProperties.FirstRow == nil ||
		*data.TableProperties.FirstRow != 1 {
		t.Fatal("the first row flag of the table properties is not set")
	}
	if data.TableProperties.TableStyleID == nil ||
		data.TableProperties.TableStyleID.StyleID != defaultTableStyleID {
		t.Fatal("the default table style of PowerPoint is not set")
	}

	// the table of the slide by the non-existing table id
	if _, err := f.GetTable(slideID, 999); !errors.As(err, new(ErrTableNotExist)) {
		t.Fatalf("expected ErrTableNotExist, got %v", err)
	}
	if err := f.SetTableCellText(slideID, 999, 0, 0, "text"); !errors.As(err, new(ErrTableNotExist)) {
		t.Fatalf("expected ErrTableNotExist, got %v", err)
	}
	// the cell out of the bounds of the table
	if err := f.SetTableCellText(slideID, tableID, 2, 0, "text"); !errors.As(err, new(ErrTableCellInvalid)) {
		t.Fatalf("expected ErrTableCellInvalid, got %v", err)
	}
}

// TestSetTableCell checks that SetTableCell and SetTableCellText set the text
// of the cell and that the text is preserved on save and on the next load.
func TestSetTableCell(t *testing.T) {
	f := NewFile()
	defer f.Close()
	slideID := f.GetSlideList()[0]

	tableID, err := f.AddTable(slideID, &TableOptions{Rows: 2, Cols: 2})
	if err != nil {
		t.Fatal(err)
	}

	if err := f.SetTableCellText(slideID, tableID, 0, 0, "plain"); err != nil {
		t.Fatal(err)
	}
	bold := 1
	size := 1800
	arial := &Font{Typeface: "Arial"}
	if err := f.SetTableCell(slideID, tableID, 1, 1, DecodeTextBody{
		BodyProperties: &DecodeBodyProperties{},
		Paragraph: []DecodeParagraph{{
			Runs: []DecodeRuns{{
				RunProperties: &DecodeRunProperties{
					Bold:  &bold,
					Size:  &size,
					Latin: arial,
				},
				Text: "styled",
			}},
		}},
	}); err != nil {
		t.Fatal(err)
	}

	table, err := f.GetTable(slideID, tableID)
	if err != nil {
		t.Fatal(err)
	}
	if got := cellText(t, table, 0, 0); got != "plain" {
		t.Fatalf("the cell (0, 0) = %q, want \"plain\"", got)
	}
	if got := cellText(t, table, 1, 1); got != "styled" {
		t.Fatalf("the cell (1, 1) = %q, want \"styled\"", got)
	}

	// save and read back: the table text and the run properties are preserved
	out := filepath.Join(t.TempDir(), "table.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	f2, err := OpenFile(out)
	if err != nil {
		t.Fatal(err)
	}
	defer f2.Close()

	table2, err := f2.GetTable(slideID, tableID)
	if err != nil {
		t.Fatal(err)
	}
	if got := cellText(t, table2, 0, 0); got != "plain" {
		t.Fatalf("after the reload the cell (0, 0) = %q, want \"plain\"", got)
	}
	run := table2.TableData.TableRows[1].TableCells[1].TextBody.Paragraph[0].Runs[0]
	if run.Text != "styled" || run.RunProperties == nil ||
		run.RunProperties.Bold == nil || *run.RunProperties.Bold != bold ||
		run.RunProperties.Size == nil || *run.RunProperties.Size != size ||
		run.RunProperties.Latin == nil || run.RunProperties.Latin.Typeface != "Arial" {
		t.Fatalf("after the reload the run of the cell (1, 1) is not preserved: %+v", run)
	}
}

// TestDeleteTable checks that DeleteTable removes the table from the slide
// and that the rest of the slide is not touched.
func TestDeleteTable(t *testing.T) {
	f := NewFile()
	defer f.Close()
	slideID := f.GetSlideList()[0]

	tableID, err := f.AddTable(slideID, &TableOptions{Rows: 2, Cols: 2})
	if err != nil {
		t.Fatal(err)
	}
	if err := f.DeleteTable(slideID, tableID); err != nil {
		t.Fatal(err)
	}
	tables, err := f.GetTables(slideID)
	if err != nil {
		t.Fatal(err)
	}
	if len(tables) != 0 {
		t.Fatalf("expected no tables, got %d", len(tables))
	}
	if err := f.DeleteTable(slideID, tableID); !errors.As(err, new(ErrTableNotExist)) {
		t.Fatalf("expected ErrTableNotExist, got %v", err)
	}

	// the rest of the slide is not touched: the shape order has no trace
	// of the deleted table
	shapeOrder, err := f.GetShapeOrder(slideID)
	if err != nil {
		t.Fatal(err)
	}
	for _, entry := range shapeOrder {
		if entry.Kind == ShapeOrderGraphicFrame {
			t.Fatalf("the shape order still contains the deleted table: %v", shapeOrder)
		}
	}
}

// TestTableRoundTrip checks that the presentation with the table that was
// added and modified with the public API keeps the table after the save and
// the next load.
func TestTableRoundTrip(t *testing.T) {
	f := NewFile()
	defer f.Close()
	slideID := f.GetSlideList()[0]

	tableID, err := f.AddTable(slideID, &TableOptions{Rows: 2, Cols: 2})
	if err != nil {
		t.Fatal(err)
	}
	if err := f.SetTableCellText(slideID, tableID, 0, 0, "a"); err != nil {
		t.Fatal(err)
	}
	if err := f.SetTableCellText(slideID, tableID, 1, 1, "b"); err != nil {
		t.Fatal(err)
	}

	out := filepath.Join(t.TempDir(), "roundtrip.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}

	f2, err := OpenFile(out)
	if err != nil {
		t.Fatal(err)
	}
	defer f2.Close()

	table, err := f2.GetTable(slideID, tableID)
	if err != nil {
		t.Fatal(err)
	}
	if got := cellText(t, table, 0, 0); got != "a" {
		t.Fatalf("the cell (0, 0) = %q, want \"a\"", got)
	}
	if got := cellText(t, table, 1, 1); got != "b" {
		t.Fatalf("the cell (1, 1) = %q, want \"b\"", got)
	}
}

// TestLoadedTableRoundTrip checks that the table of the real presentation
// (the fixture with the table created by python-pptx) is preserved byte for
// byte on the round trip: the graphic frame is not modified, so the raw
// content of the graphic data is written back as-is.
func TestLoadedTableRoundTrip(t *testing.T) {
	f, err := OpenFile(filepath.Join("test", "table.pptx"))
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()

	// the table of the fixture is read back
	tables, err := f.GetTables(256)
	if err != nil {
		t.Fatal(err)
	}
	if len(tables) != 1 {
		t.Fatalf("expected 1 table in the fixture, got %d", len(tables))
	}
	if got := cellText(t, tables[0], 0, 0); got != "Header 1" {
		t.Fatalf("the cell (0, 0) of the fixture = %q, want \"Header 1\"", got)
	}
	if got := cellText(t, tables[0], 2, 2); got != "cell 2,2" {
		t.Fatalf("the cell (2, 2) of the fixture = %q, want \"cell 2,2\"", got)
	}

	out := filepath.Join(t.TempDir(), "roundtrip.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}

	// the slide part that contains the table must be byte for byte identical
	original := readZipPart(t, filepath.Join("test", "table.pptx"), "ppt/slides/slide1.xml")
	saved := readZipPart(t, out, "ppt/slides/slide1.xml")
	if string(original) != string(saved) {
		t.Fatal("the slide part with the table is not preserved byte for byte on the round trip")
	}

	// and the table is read back from the saved file
	f2, err := OpenFile(out)
	if err != nil {
		t.Fatal(err)
	}
	defer f2.Close()
	tables2, err := f2.GetTables(256)
	if err != nil {
		t.Fatal(err)
	}
	if len(tables2) != 1 || cellText(t, tables2[0], 0, 0) != "Header 1" {
		t.Fatal("the table is not preserved on the round trip")
	}
}
