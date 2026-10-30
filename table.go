// Copyright 2026 kenny-not-dead. All rights reserved.
// Use of this source code is governed by a BSD-style license
// that can be found in the LICENSE file.
//
// Package gopptx provides functionality to create and manipulate PowerPoint
// (.pptx) files in Go, using the Office Open XML (ECMA-376) format.

package gopptx

import (
	"fmt"
)

// defaultTableStyleID is the id of the default table style of PowerPoint
// ("Medium Style 2 - Accent 1"): the accent fill of the first row and the
// banded rows.
const defaultTableStyleID = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"

// TableOptions can be passed to AddTable to set the number of rows and
// columns, the name, the position and the size of the new table. By default
// (the zero values of the position and the size) the table is centered on
// the slide, taking 80% of the slide size.
type TableOptions struct {
	// Rows is the number of rows of the table, required.
	Rows int
	// Cols is the number of columns of the table, required.
	Cols int
	// Name is the name of the graphic frame (the name attribute of the
	// p:cNvPr element), by default "Table <id>".
	Name string
	// OffsetX, OffsetY are the x and the y offsets of the table in EMU.
	OffsetX int
	OffsetY int
	// Width, Height are the width and the height of the table in EMU.
	Width  int
	Height int
}

// AddTable provides a function to add a table to the slide by the given
// slide id and returns the id of the table in the slide after it is
// appended. The table is created in the graphic frame (p:graphicFrame) with
// the default table style of PowerPoint, the columns and the rows have
// equal widths and heights, and the cells are empty. The new graphic frame
// is appended to the end of the shape tree: it is drawn above the existing
// shapes of the slide.
//
//	tableID, err := f.AddTable(slideID, &gopptx.TableOptions{
//		Rows: 3,
//		Cols: 3,
//	})
func (f *File) AddTable(slideID int, opts *TableOptions) (int, error) {
	slide, err := f.slideReader(slideID)
	if err != nil {
		return -1, err
	}

	var o TableOptions
	if opts != nil {
		o = *opts
	}
	if o.Rows <= 0 || o.Cols <= 0 {
		return -1, ErrTableInvalidSize
	}

	// the table id must be unique in the whole shape tree (sp + pic +
	// cxnSp + graphicFrame), the same way AddImage and CreateShape allocate
	// shape ids
	tableID := f.nextShapeID(slide)

	offset, extents := f.tableLayout(o)

	width := extents.CX / o.Cols
	height := extents.CY / o.Rows
	one := 1
	table := &decodeTable{
		TableProperties: &decodeTableProperties{
			FirstRow:     &one,
			BandRow:      &one,
			TableStyleID: &TableStyleID{StyleID: defaultTableStyleID},
		},
		TableGrid: &decodeTableGrid{GridCols: make([]decodeGridCol, o.Cols)},
		TableRows: make([]decodeTableRow, o.Rows),
	}
	for c := range table.TableGrid.GridCols {
		table.TableGrid.GridCols[c] = decodeGridCol{Width: &width}
	}
	for r := range table.TableRows {
		row := &table.TableRows[r]
		row.Height = &height
		row.TableCells = make([]decodeTableCell, o.Cols)
		for c := range row.TableCells {
			cell := &row.TableCells[c]
			anchor := "ctr"
			cell.TextBody = &DecodeTextBody{
				BodyProperties: &DecodeBodyProperties{Anchor: &anchor},
				ListStyle:      &struct{}{},
				Paragraph: []DecodeParagraph{{
					EndParagraphRunProperties: &DecodeRunProperties{Lang: "en-US"},
				}},
			}
			cell.TableCellProperties = &innerXML{}
		}
	}

	name := o.Name
	if name == "" {
		name = fmt.Sprintf("Table %d", tableID)
	}

	graphicFrame := decodeGraphicFrame{
		NonVisualGraphicFrameProperties: &decodeNonVisualGraphicFrameProperties{
			CommonNonVisualProperties: &decodeCommonNonVisualProperties{
				ID:   tableID,
				Name: name,
			},
			// the graphic frame locks element and the non-visual properties
			// element are written the same way PowerPoint writes them
			CommonNonVisualGraphicFrameProperties: &innerXML{Content: `<a:graphicFrameLocks noGrp="1"/>`},
			NonVisualProperties:                   &innerXML{},
		},
		Xfrm: &DecodeXfrm{
			Offset:  &offset,
			Extents: &extents,
		},
		Graphic: &decodeGraphic{
			GraphicData: &decodeGraphicData{
				URI:           graphicDataURITable,
				Table:         table,
				TableModified: true,
			},
		},
	}

	// the new graphic frame is appended to the end of the shape tree: it is
	// drawn above the existing shapes of the slide
	slide.CommonSlideData.ShapeTree.Children = append(
		slide.CommonSlideData.ShapeTree.Children,
		decodeShapeTreeChild{GraphicFrame: &graphicFrame},
	)

	return tableID, nil
}

// GetTables provides a function to get the tables of the slide by the given
// slide id, in the document order of the shape tree.
func (f *File) GetTables(slideID int) ([]Table, error) {
	var tables []Table
	slide, err := f.slideReader(slideID)
	if err != nil {
		return tables, err
	}

	for i := range slide.CommonSlideData.ShapeTree.Children {
		child := &slide.CommonSlideData.ShapeTree.Children[i]
		if child.GraphicFrame == nil {
			continue
		}
		if _, ok := child.GraphicFrame.tableData(); !ok {
			continue
		}
		tables = append(tables, newTable(*child.GraphicFrame))
	}

	return tables, nil
}

// GetTable provides a function to get the table of the slide by the given
// slide id and table id (the id attribute of the p:cNvPr element of the
// graphic frame that contains the table).
func (f *File) GetTable(slideID, tableID int) (Table, error) {
	gf, err := f.tableGraphicFrame(slideID, tableID)
	if err != nil {
		return Table{}, err
	}
	return newTable(*gf), nil
}

// SetTableCell provides a function to set the text body of the cell of the
// table by the given slide id, table id, the row and the column of the cell
// (both start from 0).
func (f *File) SetTableCell(slideID, tableID, row, col int, textBody DecodeTextBody) error {
	gf, err := f.tableGraphicFrame(slideID, tableID)
	if err != nil {
		return err
	}
	table, ok := gf.tableData()
	if !ok {
		return ErrTableNotExist{tableID}
	}
	if row < 0 || row >= len(table.TableRows) ||
		col < 0 || col >= len(table.TableRows[row].TableCells) {
		return ErrTableCellInvalid{row, col}
	}

	if textBody.BodyProperties == nil {
		textBody.BodyProperties = &DecodeBodyProperties{}
	}
	table.TableRows[row].TableCells[col].TextBody = &textBody
	gf.markTableModified()

	return nil
}

// SetTableCellText provides a function to set the text of the cell of the
// table by the given slide id, table id, the row and the column of the cell
// (both start from 0): the cell text is replaced by a single paragraph with
// a single run, keeping the body properties and the list style of the cell.
func (f *File) SetTableCellText(slideID, tableID, row, col int, text string) error {
	gf, err := f.tableGraphicFrame(slideID, tableID)
	if err != nil {
		return err
	}
	table, ok := gf.tableData()
	if !ok {
		return ErrTableNotExist{tableID}
	}
	if row < 0 || row >= len(table.TableRows) ||
		col < 0 || col >= len(table.TableRows[row].TableCells) {
		return ErrTableCellInvalid{row, col}
	}

	cell := &table.TableRows[row].TableCells[col]
	var (
		bodyProps *DecodeBodyProperties
		listStyle *struct{}
	)
	if cell.TextBody != nil {
		bodyProps = cell.TextBody.BodyProperties
		listStyle = cell.TextBody.ListStyle
	}
	if bodyProps == nil {
		bodyProps = &DecodeBodyProperties{}
	}
	if listStyle == nil {
		listStyle = &struct{}{}
	}
	cell.TextBody = &DecodeTextBody{
		BodyProperties: bodyProps,
		ListStyle:      listStyle,
		Paragraph: []DecodeParagraph{{
			Runs: []DecodeRuns{{Text: text}},
		}},
	}
	gf.markTableModified()

	return nil
}

// DeleteTable provides a function to delete the table from the slide by the
// given slide id and table id.
func (f *File) DeleteTable(slideID, tableID int) error {
	slide, err := f.slideReader(slideID)
	if err != nil {
		return err
	}

	children := slide.CommonSlideData.ShapeTree.Children

	deleteIndex := -1
	for i := range children {
		child := &children[i]
		if child.GraphicFrame == nil ||
			child.GraphicFrame.NonVisualGraphicFrameProperties == nil ||
			child.GraphicFrame.NonVisualGraphicFrameProperties.CommonNonVisualProperties == nil {
			continue
		}
		if child.GraphicFrame.NonVisualGraphicFrameProperties.CommonNonVisualProperties.ID == tableID {
			deleteIndex = i
			break
		}
	}

	if deleteIndex == -1 {
		return ErrTableNotExist{tableID}
	}

	slide.CommonSlideData.ShapeTree.Children = append(children[:deleteIndex], children[deleteIndex+1:]...)

	return nil
}

// getGraphicFrames returns the graphic frames of the slide, in the document
// order.
func (ds *decodeSlide) getGraphicFrames() []decodeGraphicFrame {
	var frames []decodeGraphicFrame
	for i := range ds.CommonSlideData.ShapeTree.Children {
		if child := ds.CommonSlideData.ShapeTree.Children[i].GraphicFrame; child != nil {
			frames = append(frames, *child)
		}
	}

	return frames
}

// tableGraphicFrame provides a function to get the graphic frame that
// contains the table by the given slide id and table id (the id attribute of
// the p:cNvPr element of the graphic frame).
func (f *File) tableGraphicFrame(slideID, tableID int) (*decodeGraphicFrame, error) {
	slide, err := f.slideReader(slideID)
	if err != nil {
		return nil, err
	}

	for i := range slide.CommonSlideData.ShapeTree.Children {
		child := &slide.CommonSlideData.ShapeTree.Children[i]
		if child.GraphicFrame == nil ||
			child.GraphicFrame.NonVisualGraphicFrameProperties == nil ||
			child.GraphicFrame.NonVisualGraphicFrameProperties.CommonNonVisualProperties == nil {
			continue
		}
		if child.GraphicFrame.NonVisualGraphicFrameProperties.CommonNonVisualProperties.ID == tableID {
			return child.GraphicFrame, nil
		}
	}

	return nil, ErrTableNotExist{tableID}
}

// tableData returns the table data (the a:tbl element of the graphic data)
// of the graphic frame, or false when the graphic frame has no table.
func (gf *decodeGraphicFrame) tableData() (*decodeTable, bool) {
	if gf.Graphic == nil || gf.Graphic.GraphicData == nil || gf.Graphic.GraphicData.Table == nil {
		return nil, false
	}
	return gf.Graphic.GraphicData.Table, true
}

// markTableModified sets the flag on the graphic data of the graphic frame,
// so the table is serialized from the model on save (instead of the raw
// content captured on load).
func (gf *decodeGraphicFrame) markTableModified() {
	gf.Graphic.GraphicData.TableModified = true
}

// newTable provides a function to convert the decode graphic frame that
// contains the table to the table.
func newTable(dg decodeGraphicFrame) Table {
	table := Table{}
	if dg.NonVisualGraphicFrameProperties != nil &&
		dg.NonVisualGraphicFrameProperties.CommonNonVisualProperties != nil {
		table.ID = dg.NonVisualGraphicFrameProperties.CommonNonVisualProperties.ID
		table.Name = dg.NonVisualGraphicFrameProperties.CommonNonVisualProperties.Name
	}
	table.Xfrm = newXfrm(dg.Xfrm)
	if data, ok := dg.tableData(); ok {
		table.TableData = newTableData(data)
	}

	return table
}

// tableLayout provides a function to calculate the offset and the extents
// (in EMU) of the new table. By default (the zero values of the options)
// the table is centered on the slide, taking 80% of the slide size, the
// same way pictureLayout does it for the pictures.
func (f *File) tableLayout(o TableOptions) (Offset, Extents) {
	slideSize := slideSize{CY: 5143500, CX: 9144000}
	if presentation, _ := f.presentationReader(); presentation != nil && presentation.SlideSize != nil {
		slideSize = *presentation.SlideSize
	}

	offset := Offset{X: slideSize.CX / 10, Y: slideSize.CY / 10}
	extents := Extents{CX: slideSize.CX * 4 / 5, CY: slideSize.CY * 4 / 5}
	if o.OffsetX != 0 {
		offset.X = o.OffsetX
	}
	if o.OffsetY != 0 {
		offset.Y = o.OffsetY
	}
	if o.Width != 0 {
		extents.CX = o.Width
	}
	if o.Height != 0 {
		extents.CY = o.Height
	}

	return offset, extents
}
