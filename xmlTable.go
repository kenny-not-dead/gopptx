// Copyright 2026 kenny-not-dead. All rights reserved.
// Use of this source code is governed by a BSD-style license
// that can be found in the LICENSE file.
//
// Package gopptx provides functionality to create and manipulate PowerPoint
// (.pptx) files in Go, using the Office Open XML (ECMA-376) format.

package gopptx

import (
	"encoding/xml"
	"regexp"
)

// graphicDataURITable defines the value of the uri attribute of the
// a:graphicData element that contains a table (a:tbl).
const graphicDataURITable = "http://schemas.openxmlformats.org/drawingml/2006/table"

// decodeGraphicFrame directly maps the p:graphicFrame element of the shape
// tree: the graphic frame of the slide that contains the table (a:tbl) or
// the other graphic content (for example the chart).
type decodeGraphicFrame struct {
	NonVisualGraphicFrameProperties *decodeNonVisualGraphicFrameProperties `xml:"nvGraphicFramePr"`
	Xfrm                            *DecodeXfrm                            `xml:"xfrm"`
	Graphic                         *decodeGraphic                         `xml:"graphic"`
}

// decodeNonVisualGraphicFrameProperties directly maps the nvGraphicFramePr
// element of the graphic frame. The cNvGraphicFramePr element (the
// a:graphicFrameLocks element) and the nvPr element are kept as-is (raw),
// the same way the other content that is not modeled is kept.
type decodeNonVisualGraphicFrameProperties struct {
	CommonNonVisualProperties             *decodeCommonNonVisualProperties `xml:"cNvPr"`
	CommonNonVisualGraphicFrameProperties *innerXML                        `xml:"cNvGraphicFramePr,omitempty"`
	NonVisualProperties                   *innerXML                        `xml:"nvPr,omitempty"`
}

type decodeGraphic struct {
	GraphicData *decodeGraphicData `xml:"graphicData"`
}

// decodeGraphicData directly maps the graphicData element of the graphic.
// The raw content of the element is captured on load and is written back
// as-is on save when the table is not modified, so the graphic data content
// that is not modeled (for example the chart) is preserved on the round trip.
type decodeGraphicData struct {
	URI   string `xml:"uri,attr"`
	Table *decodeTable

	// RawContent is the raw content of the element as it is in the slide
	// XML; it is empty for the graphic frames created with AddTable.
	RawContent string `xml:"-"`
	// TableModified is set when the table is modified with the public API
	// (SetTableCell, SetTableCellText): the table is then serialized from
	// the model on save instead of the raw content.
	TableModified bool `xml:"-"`
}

// tblElementRe matches the table element (tbl, with any namespace prefix) in
// the raw content of the graphic data element.
var tblElementRe = regexp.MustCompile(`<(?:[A-Za-z_][\w.-]*:)?tbl[ >/]`)

// UnmarshalXML reads the graphic data element: the uri attribute, the table
// (a:tbl) when it is present, and the raw content of the element.
func (dst *decodeGraphicData) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		if attr.Name.Local == "uri" {
			dst.URI = attr.Value
		}
	}

	raw := new(innerXML)
	if err := d.DecodeElement(raw, &start); err != nil {
		return err
	}
	dst.RawContent = raw.Content

	if !tblElementRe.MatchString(raw.Content) {
		return nil
	}

	// parse the table from the raw content: the element names use the
	// prefixes of the slide document, the struct tags match by the local
	// name only
	var wrapped struct {
		Table *decodeTable `xml:"tbl"`
	}
	if err := xml.Unmarshal([]byte("<wrap>"+raw.Content+"</wrap>"), &wrapped); err != nil {
		return err
	}
	dst.Table = wrapped.Table

	return nil
}

// decodeTable directly maps the tbl element of the graphic data: the table
// properties, the grid (the widths of the columns) and the rows.
type decodeTable struct {
	TableProperties *decodeTableProperties `xml:"tblPr"`
	TableGrid       *decodeTableGrid       `xml:"tblGrid"`
	TableRows       []decodeTableRow       `xml:"tr"`
}

type decodeTableProperties struct {
	FirstRow     *int          `xml:"firstRow,attr,omitempty"`
	BandRow      *int          `xml:"bandRow,attr,omitempty"`
	TableStyleID *TableStyleID `xml:"tableStyleId,omitempty"`
}

type decodeTableGrid struct {
	GridCols []decodeGridCol `xml:"gridCol"`
}

type decodeGridCol struct {
	Width *int `xml:"w,attr"`
}

type decodeTableRow struct {
	Height     *int              `xml:"h,attr,omitempty"`
	TableCells []decodeTableCell `xml:"tc"`
}

// decodeTableCell directly maps the tc element of the table row. The tcPr
// element (the cell properties: the borders, the fill and the margins) is
// kept as-is (raw).
type decodeTableCell struct {
	GridSpan            *int            `xml:"gridSpan,attr,omitempty"`
	RowSpan             *int            `xml:"rowSpan,attr,omitempty"`
	HMerge              *bool           `xml:"hMerge,attr,omitempty"`
	VMerge              *bool           `xml:"vMerge,attr,omitempty"`
	TextBody            *DecodeTextBody `xml:"txBody"`
	TableCellProperties *innerXML       `xml:"tcPr,omitempty"`
}

// GraphicFrame directly maps the p:graphicFrame element of the shape tree:
// the graphic frame of the slide that contains the table (a:tbl).
type GraphicFrame struct {
	NonVisualGraphicFrameProperties *NonVisualGraphicFrameProperties `xml:"p:nvGraphicFramePr"`
	Xfrm                            *Xfrm                            `xml:"p:xfrm"`
	Graphic                         *Graphic                         `xml:"a:graphic"`
}

// NonVisualGraphicFrameProperties defines the non-visual properties of the
// graphic frame (p:nvGraphicFramePr).
type NonVisualGraphicFrameProperties struct {
	CommonNonVisualProperties             *CommonNonVisualProperties `xml:"p:cNvPr"`
	CommonNonVisualGraphicFrameProperties *innerXML                  `xml:"p:cNvGraphicFramePr,omitempty"`
	NonVisualProperties                   *innerXML                  `xml:"p:nvPr,omitempty"`
}

// Graphic directly maps the a:graphic element of the graphic frame.
type Graphic struct {
	GraphicData *GraphicData `xml:"a:graphicData"`
}

// GraphicData directly maps the a:graphicData element of the graphic.
type GraphicData struct {
	URI       string     `xml:"-"`
	TableData *TableData `xml:"a:tbl"`

	// RawContent is the raw content of the element; on save it is written
	// back as-is when the table is not modified, so the graphic data
	// content that is not modeled (for example the chart) is preserved on
	// the round trip.
	RawContent    string `xml:"-"`
	TableModified bool   `xml:"-"`
}

// rawGraphicData is the raw form of the graphic data element: the raw
// content is written as-is.
type rawGraphicData struct {
	URI     string `xml:"uri,attr"`
	Content string `xml:",innerxml"`
}

// MarshalXML writes the graphic data element: the table from the model when
// the table is modified, the raw content as-is otherwise.
func (g GraphicData) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = xml.Name{Local: "a:graphicData"}
	if g.TableData == nil || !g.TableModified {
		return e.EncodeElement(rawGraphicData{URI: g.URI, Content: g.RawContent}, start)
	}
	start.Attr = []xml.Attr{{Name: xml.Name{Local: "uri"}, Value: g.URI}}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if err := e.EncodeElement(g.TableData, xml.StartElement{Name: xml.Name{Local: "a:tbl"}}); err != nil {
		return err
	}
	return e.EncodeToken(xml.EndElement{Name: start.Name})
}

// TableData directly maps the a:tbl element of the graphic data: the table
// properties, the grid (the widths of the columns) and the rows.
type TableData struct {
	TableProperties *TableProperties `xml:"a:tblPr"`
	TableGrid       *TableGrid       `xml:"a:tblGrid"`
	TableRows       []TableRow       `xml:"a:tr"`
}

// TableProperties defines the table properties (a:tblPr) element: the flags
// of the first row and the banded rows, and the id of the table style.
type TableProperties struct {
	FirstRow     *int          `xml:"firstRow,attr,omitempty"`
	BandRow      *int          `xml:"bandRow,attr,omitempty"`
	TableStyleID *TableStyleID `xml:"a:tableStyleId,omitempty"`
}

// TableStyleID defines the a:tableStyleId element of the table properties:
// the id of the built-in table style of PowerPoint.
type TableStyleID struct {
	StyleID string `xml:",chardata"`
}

// TableGrid defines the table grid (a:tblGrid) element: the widths of the
// columns.
type TableGrid struct {
	GridCols []GridCol `xml:"a:gridCol"`
}

// GridCol defines a column of the table grid (a:gridCol) element.
type GridCol struct {
	Width *int `xml:"w,attr"`
}

// TableRow defines a row of the table (a:tr) element.
type TableRow struct {
	Height     *int        `xml:"h,attr,omitempty"`
	TableCells []TableCell `xml:"a:tc"`
}

// TableCell defines a cell of the table (a:tc) element. The cell properties
// (a:tcPr: the borders, the fill and the margins) are kept as-is (raw).
type TableCell struct {
	GridSpan            *int      `xml:"gridSpan,attr,omitempty"`
	RowSpan             *int      `xml:"rowSpan,attr,omitempty"`
	HMerge              *bool     `xml:"hMerge,attr,omitempty"`
	VMerge              *bool     `xml:"vMerge,attr,omitempty"`
	TextBody            *TextBody `xml:"a:txBody"`
	TableCellProperties *innerXML `xml:"a:tcPr,omitempty"`
}

// Table is a table of the slide: the id and the name of the graphic frame
// that contains the table, the position and the size of the graphic frame,
// and the table data (the a:tbl element of the graphic data).
type Table struct {
	ID        int
	Name      string
	Xfrm      *Xfrm
	TableData *TableData
}

// newTableData provides a function to convert the decode table to the table
// data.
func newTableData(dt *decodeTable) *TableData {
	if dt == nil {
		return nil
	}

	td := &TableData{TableRows: make([]TableRow, len(dt.TableRows))}
	if dt.TableProperties != nil {
		td.TableProperties = &TableProperties{
			FirstRow:     dt.TableProperties.FirstRow,
			BandRow:      dt.TableProperties.BandRow,
			TableStyleID: dt.TableProperties.TableStyleID,
		}
	}
	if dt.TableGrid != nil {
		grid := &TableGrid{GridCols: make([]GridCol, len(dt.TableGrid.GridCols))}
		for i := range dt.TableGrid.GridCols {
			grid.GridCols[i] = GridCol{Width: dt.TableGrid.GridCols[i].Width}
		}
		td.TableGrid = grid
	}
	for i := range dt.TableRows {
		row := &dt.TableRows[i]
		td.TableRows[i] = TableRow{
			Height:     row.Height,
			TableCells: make([]TableCell, len(row.TableCells)),
		}
		for j := range row.TableCells {
			cell := &row.TableCells[j]
			td.TableRows[i].TableCells[j] = TableCell{
				GridSpan:            cell.GridSpan,
				RowSpan:             cell.RowSpan,
				HMerge:              cell.HMerge,
				VMerge:              cell.VMerge,
				TextBody:            newTextBody(cell.TextBody),
				TableCellProperties: cell.TableCellProperties,
			}
		}
	}

	return td
}

// newGraphicFrame provides a function to convert the decode graphic frame to
// the graphic frame.
func newGraphicFrame(dg decodeGraphicFrame) GraphicFrame {
	gf := GraphicFrame{}
	if dg.NonVisualGraphicFrameProperties != nil {
		gf.NonVisualGraphicFrameProperties = &NonVisualGraphicFrameProperties{
			CommonNonVisualProperties:             newCommonNonVisualProperties(dg.NonVisualGraphicFrameProperties.CommonNonVisualProperties),
			CommonNonVisualGraphicFrameProperties: dg.NonVisualGraphicFrameProperties.CommonNonVisualGraphicFrameProperties,
			NonVisualProperties:                   dg.NonVisualGraphicFrameProperties.NonVisualProperties,
		}
	}
	gf.Xfrm = newXfrm(dg.Xfrm)
	if dg.Graphic != nil && dg.Graphic.GraphicData != nil {
		gd := dg.Graphic.GraphicData
		gf.Graphic = &Graphic{
			GraphicData: &GraphicData{
				URI:           gd.URI,
				RawContent:    gd.RawContent,
				TableModified: gd.TableModified,
			},
		}
		if gd.Table != nil {
			gf.Graphic.GraphicData.TableData = newTableData(gd.Table)
		}
	}

	return gf
}
