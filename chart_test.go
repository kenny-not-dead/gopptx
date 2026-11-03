// Copyright 2026 kenny-not-dead. All rights reserved.
// Use of this source code is governed by a BSD-style license
// that can be found in the LICENSE file.
//
// Package gopptx provides functionality to create and manipulate PowerPoint
// (.pptx) files in Go, using the Office Open XML (ECMA-376) format.

package gopptx

import (
	"bytes"
	"encoding/xml"
	"errors"
	"path/filepath"
	"sort"
	"strconv"
	"strings"
	"testing"
)

// normalizeSldSz sorts the attributes of the sldSz element of the
// presentation part, so the part is compared up to the order of the
// attributes of the element.
func normalizeSldSz(content []byte) []byte {
	start := bytes.Index(content, []byte("<p:sldSz"))
	if start == -1 {
		return content
	}
	end := bytes.IndexByte(content[start+1:], '>')
	if end == -1 {
		return content
	}
	end += start + 1
	tag := content[start : end+1]
	i := bytes.IndexByte(tag, ' ')
	if i == -1 {
		return content
	}
	name := tag[:i]
	attrs := tag[i+1 : len(tag)-1] // without the final >
	closing := ">"
	if bytes.HasSuffix(attrs, []byte("/")) {
		closing = "/>"
		attrs = attrs[:len(attrs)-1]
	}
	fields := strings.Fields(string(attrs))
	sort.Strings(fields)
	var b strings.Builder
	b.WriteString(string(name))
	if len(fields) > 0 {
		b.WriteString(" ")
		b.WriteString(strings.Join(fields, " "))
	}
	b.WriteString(closing)
	return append(append(append([]byte{}, content[:start]...), b.String()...), content[end+1:]...)
}

// TestAddChart checks that AddChart creates the chart part and the graphic
// frame (the c:chart reference) in the slide, and that the chart can be read
// back.
func TestAddChart(t *testing.T) {
	f := NewFile()
	defer f.Close()
	slideID := f.GetSlideList()[0]

	// the data of the chart must have the categories and the series
	if _, err := f.AddChart(slideID, nil, nil, nil); !errors.Is(err, ErrChartInvalidData) {
		t.Fatalf("expected ErrChartInvalidData, got %v", err)
	}
	if _, err := f.AddChart(slideID, []string{"a", "b"},
		[]ChartSeries{{Name: "s", Values: []float64{1}}}, nil); !errors.Is(err, ErrChartInvalidData) {
		t.Fatalf("expected ErrChartInvalidData, got %v", err)
	}
	if _, err := f.AddChart(999, []string{"a"}, []ChartSeries{{Name: "s", Values: []float64{1}}}, nil); !errors.As(err, new(ErrSlideNotExist)) {
		t.Fatalf("expected ErrSlideNotExist, got %v", err)
	}

	chartID, err := f.AddChart(slideID,
		[]string{"2024", "2025", "2026"},
		[]ChartSeries{
			{Name: "Revenue", Values: []float64{120, 180, 250}},
			{Name: "Costs", Values: []float64{80, 95, 120}},
		},
		&ChartOptions{Title: "Revenue and costs"})
	if err != nil {
		t.Fatal(err)
	}

	charts, err := f.GetCharts(slideID)
	if err != nil {
		t.Fatal(err)
	}
	if len(charts) != 1 {
		t.Fatalf("expected 1 chart, got %d", len(charts))
	}
	if frameID(charts[0]) != chartID {
		t.Fatalf("chart id = %d, want %d", frameID(charts[0]), chartID)
	}
	if frameName(charts[0]) != "Chart "+strconv.Itoa(chartID) {
		t.Fatalf("chart name = %q, want the default name", frameName(charts[0]))
	}
	if chartRefID(charts[0]) == "" {
		t.Fatal("the chart has no relationship id")
	}

	chart, err := f.GetChart(slideID, chartID)
	if err != nil {
		t.Fatal(err)
	}
	if frameID(chart) != chartID || chartRefID(chart) == "" {
		t.Fatalf("the chart is not read back: %+v", chart)
	}
	// the chart of the slide by the non-existing chart id
	if _, err := f.GetChart(slideID, 999); !errors.As(err, new(ErrChartNotExist)) {
		t.Fatalf("expected ErrChartNotExist, got %v", err)
	}

	out := filepath.Join(t.TempDir(), "chart.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}

	// the graphic frame of the slide with the c:chart reference
	slideXML := string(readZipPart(t, out, "ppt/slides/slide1.xml"))
	if !strings.Contains(slideXML, `<p:graphicFrame>`) {
		t.Fatalf("the graphic frame of the chart is not written in the slide XML:\n%s", slideXML)
	}
	if !strings.Contains(slideXML, `uri="`+graphicDataURIChart+`"`) {
		t.Fatalf("the graphic data of the chart is not written in the slide XML:\n%s", slideXML)
	}
	if !strings.Contains(slideXML, `<c:chart xmlns:c="`+NameSpaceDrawingMLChart.Value+`"`) {
		t.Fatalf("the c:chart reference with the inline namespace is not written in the slide XML:\n%s", slideXML)
	}
	if !strings.Contains(slideXML, `r:id="`+chartRefID(chart)+`"`) {
		t.Fatalf("the chart reference does not use the relationship id %s in the slide XML:\n%s", chartRefID(chart), slideXML)
	}

	// the chart relationship in the slide rels part
	rels := string(readZipPart(t, out, "ppt/slides/_rels/slide1.xml.rels"))
	if !strings.Contains(rels, SourceRelationshipChart) || !strings.Contains(rels, "../charts/chart1.xml") {
		t.Fatalf("the chart relationship is not written in the slide rels part:\n%s", rels)
	}

	// the chart part with the data of the series in the cache
	chartXML := string(readZipPart(t, out, "ppt/charts/chart1.xml"))
	for _, want := range []string{
		"<c:barChart>", `<c:barDir val="col"/>`, `Revenue`, `Costs`,
		`2024`, `120`, `250`, `Revenue and costs`,
	} {
		if !strings.Contains(chartXML, want) {
			t.Fatalf("the chart part does not contain %q:\n%s", want, chartXML)
		}
	}
	if strings.Contains(chartXML, "externalData") {
		t.Fatalf("the chart part must not reference the external data:\n%s", chartXML)
	}

	// the chart part mirrors the structure of the chart part that
	// PowerPoint writes: the language and the rounded corners of the chart
	// space, the explicit text properties of the chart (the font size of
	// the text that has no size of its own) and of the axes and the
	// legend, the numeric format of the axes, and the positive ids of the
	// axes
	for _, want := range []string{
		`<c:lang val="en-US"/>`,
		`<c:roundedCorners val="0"/>`,
		`<c:varyColors val="0"/>`,
		`<c:gapWidth val="150"/>`,
		`<c:overlap val="0"/>`,
		`<c:plotVisOnly val="1"/>`,
		`<c:showDLblsOverMax val="0"/>`,
		`<c:numFmt formatCode="General" sourceLinked="1"/>`,
		`<a:defRPr sz="1800"/>`,
		`<a:defRPr sz="1200"/>`,
	} {
		if !strings.Contains(chartXML, want) {
			t.Fatalf("the chart part does not contain %q:\n%s", want, chartXML)
		}
	}
	if strings.Contains(chartXML, `val="-`) {
		t.Fatalf("the chart part has the negative ids of the axes:\n%s", chartXML)
	}
	// the chart part is the well-formed XML
	{
		var generic interface{}
		if err := xml.Unmarshal([]byte(chartXML), &generic); err != nil {
			t.Fatalf("the chart part is not the well-formed XML: %v\n%s", err, chartXML)
		}
	}

	// the content type of the chart part
	contentTypes := string(readZipPart(t, out, "[Content_Types].xml"))
	if !strings.Contains(contentTypes, ContentTypeChartML) {
		t.Fatalf("the content type of the chart part is not written:\n%s", contentTypes)
	}
}

// TestChartRoundTrip checks that the presentation with the chart that was
// added and the table that was modified with the public API keeps both after
// the save and the next load.
func TestChartRoundTrip(t *testing.T) {
	f := NewFile()
	defer f.Close()
	slideID := f.GetSlideList()[0]

	chartID, err := f.AddChart(slideID,
		[]string{"a", "b"},
		[]ChartSeries{{Name: "s", Values: []float64{1, 2}}},
		nil)
	if err != nil {
		t.Fatal(err)
	}
	tableID, err := f.AddTable(slideID, &TableOptions{Rows: 2, Cols: 2})
	if err != nil {
		t.Fatal(err)
	}
	if err := f.SetTableCellText(slideID, tableID, 0, 0, "a"); err != nil {
		t.Fatal(err)
	}

	out := filepath.Join(t.TempDir(), "roundtrip.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	chartPartBefore := readZipPart(t, out, "ppt/charts/chart1.xml")

	f2, err := OpenFile(out)
	if err != nil {
		t.Fatal(err)
	}
	defer f2.Close()

	chart, err := f2.GetChart(slideID, chartID)
	if err != nil {
		t.Fatal(err)
	}
	if chartRefID(chart) == "" {
		t.Fatal("the chart is not read back")
	}
	table, err := f2.GetTable(slideID, tableID)
	if err != nil {
		t.Fatal(err)
	}
	if got := cellText(t, table, 0, 0); got != "a" {
		t.Fatalf("the cell (0, 0) = %q, want \"a\"", got)
	}

	// the chart part is preserved on the round trip
	chartPartAfter := readZipPart(t, out, "ppt/charts/chart1.xml")
	if string(chartPartBefore) != string(chartPartAfter) {
		t.Fatal("the chart part differs after the round trip")
	}
}

// TestLoadedChartRoundTrip checks that the chart of the real presentation
// (the fixture with the chart created by python-pptx) is preserved on the
// round trip: the graphic frame is not modified, so the raw content of the
// graphic data (the c:chart reference) is written back as-is, and the chart
// part and its embedded workbook are kept byte for byte.
func TestLoadedChartRoundTrip(t *testing.T) {
	fixture := filepath.Join("test", "chart.pptx")
	f, err := OpenFile(fixture)
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()

	// the chart of the fixture is read back
	charts, err := f.GetCharts(256)
	if err != nil {
		t.Fatal(err)
	}
	if len(charts) != 1 {
		t.Fatalf("expected 1 chart in the fixture, got %d", len(charts))
	}
	if frameName(charts[0]) != "Chart 2" {
		t.Fatalf("the chart name = %q, want \"Chart 2\"", frameName(charts[0]))
	}
	if chartRefID(charts[0]) != "rId2" {
		t.Fatalf("the chart relationship id = %q, want \"rId2\"", chartRefID(charts[0]))
	}

	out := filepath.Join(t.TempDir(), "roundtrip.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	// every part is preserved byte for byte, the presentation part is
	// compared up to the order of the attributes of the sldSz element (the
	// known attribute order difference of the element)
	if err := assertPartsEqualWith(t, fixture, out, func(name string, content []byte) []byte {
		if name == "ppt/presentation.xml" {
			return normalizeSldSz(content)
		}
		return content
	}); err != nil {
		t.Error(err)
	}

	// and the chart is read back from the saved file
	f2, err := OpenFile(out)
	if err != nil {
		t.Fatal(err)
	}
	defer f2.Close()
	charts2, err := f2.GetCharts(256)
	if err != nil {
		t.Fatal(err)
	}
	if len(charts2) != 1 || chartRefID(charts2[0]) != "rId2" {
		t.Fatal("the chart is not preserved on the round trip")
	}
}

// TestDeleteChart checks that DeleteChart removes the graphic frame, the
// chart relationship and the chart part (including the embedded workbook)
// from the presentation.
func TestDeleteChart(t *testing.T) {
	fixture := filepath.Join("test", "chart.pptx")
	f, err := OpenFile(fixture)
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()

	charts, err := f.GetCharts(256)
	if err != nil {
		t.Fatal(err)
	}
	if len(charts) != 1 {
		t.Fatalf("expected 1 chart in the fixture, got %d", len(charts))
	}
	if err := f.DeleteChart(256, frameID(charts[0])); err != nil {
		t.Fatal(err)
	}
	if err := f.DeleteChart(256, frameID(charts[0])); !errors.As(err, new(ErrChartNotExist)) {
		t.Fatalf("expected ErrChartNotExist, got %v", err)
	}
	// the chart of the slide by the non-existing chart id
	if _, err := f.GetChart(256, 999); !errors.As(err, new(ErrChartNotExist)) {
		t.Fatalf("expected ErrChartNotExist, got %v", err)
	}

	out := filepath.Join(t.TempDir(), "chart_deleted.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	// the graphic frame is removed from the slide
	slideXML := string(readZipPart(t, out, "ppt/slides/slide1.xml"))
	if strings.Contains(slideXML, "<p:graphicFrame>") {
		t.Fatalf("the graphic frame of the chart is not removed from the slide XML:\n%s", slideXML)
	}
	// the chart relationship is removed from the slide rels part
	rels := string(readZipPart(t, out, "ppt/slides/_rels/slide1.xml.rels"))
	if strings.Contains(rels, "relationships/chart") {
		t.Fatalf("the chart relationship is not removed from the slide rels part:\n%s", rels)
	}
	// the chart part and its embedded workbook are removed from the package
	if hasZipPart(t, out, "ppt/charts/chart1.xml") {
		t.Fatal("the chart part is not removed from the package")
	}
	if hasZipPart(t, out, "ppt/charts/_rels/chart1.xml.rels") {
		t.Fatal("the rels part of the chart part is not removed from the package")
	}
	if hasZipPart(t, out, "ppt/embeddings/Microsoft_Excel_Sheet1.xlsx") {
		t.Fatal("the embedded workbook of the chart is not removed from the package")
	}
	// the content types part keeps the origin content without the chart
	// override; the xlsx default declared in the origin content types part
	// is kept, the same way the defaults of the removed images are kept
	contentTypes := string(readZipPart(t, out, "[Content_Types].xml"))
	if strings.Contains(contentTypes, ContentTypeChartML) {
		t.Fatalf("the content type of the chart part is not removed:\n%s", contentTypes)
	}
	if !strings.Contains(contentTypes, `Extension="xlsx"`) {
		t.Fatalf("the xlsx default of the origin content types part is not kept:\n%s", contentTypes)
	}
}

// TestAddDeleteChart checks that adding and then deleting the chart keeps
// the presentation content unchanged, the same way TestAddDeleteImage does
// it for the images.
func TestAddDeleteChart(t *testing.T) {
	f, err := OpenFile(testFixture)
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()

	chartID, err := f.AddChart(257,
		[]string{"a", "b"},
		[]ChartSeries{{Name: "s", Values: []float64{1, 2}}},
		nil)
	if err != nil {
		t.Fatal(err)
	}
	if err := f.DeleteChart(257, chartID); err != nil {
		t.Fatal(err)
	}

	out := filepath.Join(t.TempDir(), "add_delete.pptx")
	if err := f.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	if err := assertPartsEqual(t, testFixture, out); err != nil {
		t.Fatal(err)
	}
}
