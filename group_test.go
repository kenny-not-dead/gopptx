// Copyright 2026 kenny-not-dead. All rights reserved.
// Use of this source code is governed by a BSD-style license
// that can be found in the LICENSE file.
//
// Package gopptx provides functionality to create and manipulate PowerPoint
// (.pptx) files in Go, using the Office Open XML (ECMA-376) format.

package gopptx

import (
	"archive/zip"
	"bytes"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/stretchr/testify/assert"
)

// groupHyperlinkTarget is the target of the hyperlink relationship of the
// shape with the id 101 (the shape is inside the group) of the group fixture.
const groupHyperlinkTarget = "https://github.com/kenny-not-dead/gopptx"

// groupShapeHyperlinkTarget is the target of the hyperlink relationship of
// the top level group (the id 100) of the group fixture.
const groupShapeHyperlinkTarget = "https://example.com"

// groupSlideInjection is the content injected into the shape tree of the
// empty slide 257 of the test/presentation.pptx fixture: the shape before
// the group (the id 105), the top level group (the id 100, with the
// hyperlink on the group non-visual properties) with the shape with the text
// and the hyperlink (the id 101), the shape without the text body (the id
// 102) and the nested group (the id 103) with the shape (the id 104), and
// the shape after the group (the id 106).
const groupSlideInjection = `<p:sp><p:nvSpPr><p:cNvPr id="105" name="Rectangle 105"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="2000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp><p:grpSp><p:nvGrpSpPr><p:cNvPr id="100" name="Group 100"><a:hlinkClick r:id="rId3"/></p:cNvPr><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="300000" y="300000"/><a:ext cx="4000000" cy="3000000"/><a:chOff x="300000" y="300000"/><a:chExt cx="4000000" cy="3000000"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="101" name="Text 101"><a:hlinkClick r:id="rId2"/></p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="500000" y="500000"/><a:ext cx="1500000" cy="500000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>inside group</a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="102" name="Rectangle 102"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="500000" y="1200000"/><a:ext cx="1500000" cy="500000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp><p:grpSp><p:nvGrpSpPr><p:cNvPr id="103" name="Group 103"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="2200000" y="500000"/><a:ext cx="1500000" cy="1200000"/><a:chOff x="2200000" y="500000"/><a:chExt cx="1500000" cy="1200000"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="104" name="Text 104"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="2400000" y="700000"/><a:ext cx="1100000" cy="500000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>inside nested group</a:t></a:r></a:p></p:txBody></p:sp></p:grpSp></p:grpSp><p:sp><p:nvSpPr><p:cNvPr id="106" name="Ellipse 106"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="7000000" y="5000000"/><a:ext cx="800000" cy="800000"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom></p:spPr></p:sp>`

// groupFixture provides a function to build the presentation fixture with
// the group shapes: the test/presentation.pptx fixture with the shape tree
// of the empty slide 257 (the ppt/slides/slide2.xml part) extended by the
// group shapes (the nested group included) and the hyperlink relationships
// of the shape with the id 101 (rId2) and of the top level group (rId3)
// added to the rels part of the slide. The fixture is written to a temporary
// file, the path of the fixture is returned.
func groupFixture(t *testing.T) string {
	t.Helper()

	src, err := os.Open(testFixture)
	assert.NoError(t, err)
	defer src.Close()
	srcInfo, err := src.Stat()
	assert.NoError(t, err)
	srcReader, err := zip.NewReader(src, srcInfo.Size())
	assert.NoError(t, err)

	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	relsInjection := `<Relationship Id="rId2" Type="` + SourceRelationshipHyperlink + `" Target="` + groupHyperlinkTarget + `" TargetMode="External"/>` +
		`<Relationship Id="rId3" Type="` + SourceRelationshipHyperlink + `" Target="` + groupShapeHyperlinkTarget + `" TargetMode="External"/>`
	for _, v := range srcReader.File {
		rc, err := v.Open()
		assert.NoError(t, err)
		var content bytes.Buffer
		_, err = content.ReadFrom(rc)
		assert.NoError(t, err)
		assert.NoError(t, rc.Close())

		switch v.Name {
		case "ppt/slides/slide2.xml":
			assert.Contains(t, string(content.Bytes()), "</p:spTree>")
			content = *bytes.NewBuffer(bytes.Replace(content.Bytes(), []byte("</p:spTree>"), []byte(groupSlideInjection+"</p:spTree>"), 1))
		case "ppt/slides/_rels/slide2.xml.rels":
			assert.Contains(t, string(content.Bytes()), "</Relationships>")
			content = *bytes.NewBuffer(bytes.Replace(content.Bytes(), []byte("</Relationships>"), []byte(relsInjection+"</Relationships>"), 1))
		}

		w, err := zw.Create(v.Name)
		assert.NoError(t, err)
		_, err = w.Write(content.Bytes())
		assert.NoError(t, err)
	}
	assert.NoError(t, zw.Close())

	path := filepath.Join(t.TempDir(), "groups.pptx")
	assert.NoError(t, os.WriteFile(path, buf.Bytes(), 0o644))

	return path
}

// TestGetGroups checks that the group shapes of the slide are read back: the
// groups (including the nested groups) of the shape tree, the shapes of the
// groups, the text of the shapes of the groups and the group properties.
func TestGetGroups(t *testing.T) {
	f, err := OpenFile(groupFixture(t))
	assert.NoError(t, err)
	defer f.Close()

	// the slide without the groups
	groups, err := f.GetGroups(256)
	assert.NoError(t, err)
	assert.Empty(t, groups)

	// the slide with the groups
	groups, err = f.GetGroups(257)
	assert.NoError(t, err)
	assert.Len(t, groups, 1)

	group := groups[0].ShapeTree
	assert.Equal(t, 100, group.NonVisualGroupShapeProperties.CommonNonVisualProperties.ID)
	assert.Equal(t, "Group 100", group.NonVisualGroupShapeProperties.CommonNonVisualProperties.Name)

	// the shapes of the group, in the document order
	assert.Len(t, group.Shape, 2)
	assert.Equal(t, 101, group.Shape[0].NonVisualShapeProperties.CommonNonVisualProperties.ID)
	assert.Equal(t, 102, group.Shape[1].NonVisualShapeProperties.CommonNonVisualProperties.ID)
	assert.Equal(t, "inside group", group.Shape[0].TextBody.Paragraph[0].Runs[0].Text)
	assert.Nil(t, group.Shape[1].TextBody)

	// the nested group of the group
	assert.Len(t, group.GroupShape, 1)
	nested := group.GroupShape[0].ShapeTree
	assert.Equal(t, 103, nested.NonVisualGroupShapeProperties.CommonNonVisualProperties.ID)
	assert.Equal(t, "Group 103", nested.NonVisualGroupShapeProperties.CommonNonVisualProperties.Name)
	assert.Len(t, nested.Shape, 1)
	assert.Equal(t, 104, nested.Shape[0].NonVisualShapeProperties.CommonNonVisualProperties.ID)
	assert.Equal(t, "inside nested group", nested.Shape[0].TextBody.Paragraph[0].Runs[0].Text)

	// the group shape properties (the xfrm of the group)
	assert.Equal(t, 300000, group.GroupShapeProperties.Xfrm.Offset.X)
	assert.Equal(t, 4000000, group.GroupShapeProperties.Xfrm.Extents.CX)
	assert.Equal(t, 300000, group.GroupShapeProperties.Xfrm.ChildOffset.X)
	assert.Equal(t, 3000000, group.GroupShapeProperties.Xfrm.ChildExtents.CY)

	// the groups of the slide are not returned by the shapes of the slide
	shapes, err := f.GetShapes(257)
	assert.NoError(t, err)
	assert.Len(t, shapes, 4)

	// the slide that is not in the presentation
	_, err = f.GetGroups(999)
	assert.EqualError(t, err, ErrSlideNotExist{SlideID: 999}.Error())
}

// TestGroupsRoundTrip checks that the group shapes (including the nested
// groups) of the slide are preserved on the round trip: the opened and the
// saved presentation have the identical parts.
func TestGroupsRoundTrip(t *testing.T) {
	fixture := groupFixture(t)

	f, err := OpenFile(fixture)
	assert.NoError(t, err)
	defer f.Close()

	// read the shapes and the groups of the slide to trigger the slide parsing
	_, err = f.GetShapes(257)
	assert.NoError(t, err)
	_, err = f.GetGroups(257)
	assert.NoError(t, err)

	out := filepath.Join(t.TempDir(), "groups.pptx")
	assert.NoError(t, f.SaveAs(out))
	assert.NoError(t, assertPartsEqual(t, fixture, out))
}

// TestGroupShapeOrder checks that the group shapes take part in the z-order
// of the shape tree of the slide: the group is in the shape order with the
// ShapeOrderGrpSp kind, the order of the group can be changed by the
// SetShapeOrder and the MoveShape functions and the new order is preserved
// on the round trip.
func TestGroupShapeOrder(t *testing.T) {
	f, err := OpenFile(groupFixture(t))
	assert.NoError(t, err)
	defer f.Close()

	// the order of the shape tree of the slide: the placeholders (the ids 7
	// and 8), the shape before the group (the id 105), the group (the id
	// 100) and the shape after the group (the id 106)
	shapeOrder, err := f.GetShapeOrder(257)
	assert.NoError(t, err)
	assert.Equal(t, []ShapeOrder{
		{Kind: ShapeOrderSp, ID: 7},
		{Kind: ShapeOrderSp, ID: 8},
		{Kind: ShapeOrderSp, ID: 105},
		{Kind: ShapeOrderGrpSp, ID: 100},
		{Kind: ShapeOrderSp, ID: 106},
	}, shapeOrder)

	// the invalid shape order: the entry of the group is missing
	invalidOrder := []ShapeOrder{
		{Kind: ShapeOrderSp, ID: 7},
		{Kind: ShapeOrderSp, ID: 8},
		{Kind: ShapeOrderSp, ID: 105},
	}
	err = f.SetShapeOrder(257, invalidOrder)
	assert.EqualError(t, err, ErrShapeOrderInvalid.Error())

	// move the group before the first shape of the slide
	err = f.MoveShape(257, 100, 7)
	assert.NoError(t, err)
	shapeOrder, err = f.GetShapeOrder(257)
	assert.NoError(t, err)
	assert.Equal(t, []ShapeOrder{
		{Kind: ShapeOrderGrpSp, ID: 100},
		{Kind: ShapeOrderSp, ID: 7},
		{Kind: ShapeOrderSp, ID: 8},
		{Kind: ShapeOrderSp, ID: 105},
		{Kind: ShapeOrderSp, ID: 106},
	}, shapeOrder)

	// the order that is not the permutation of the shape tree of the slide
	err = f.SetShapeOrder(257, []ShapeOrder{{Kind: ShapeOrderSp, ID: 7}})
	assert.EqualError(t, err, ErrShapeOrderInvalid.Error())

	// the shape (the group) that is not in the shape tree of the slide
	err = f.MoveShape(257, 999, 7)
	assert.EqualError(t, err, ErrShapeNotExist{999}.Error())

	// the order of the group shapes is preserved on the round trip
	out := filepath.Join(t.TempDir(), "groups.pptx")
	assert.NoError(t, f.SaveAs(out))
	groups := string(readZipPart(t, out, "ppt/slides/slide2.xml"))
	// the top level group and the nested group are in the saved slide
	assert.Equal(t, 2, strings.Count(groups, "<p:grpSp>"))
	groupIdx := strings.Index(groups, `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="100"`)
	shapeIdx := strings.Index(groups, `<p:sp><p:nvSpPr><p:cNvPr id="7"`)
	assert.Greater(t, shapeIdx, -1)
	assert.Less(t, groupIdx, shapeIdx)

	// the content of the group is kept after the move
	reopened, err := OpenFile(out)
	assert.NoError(t, err)
	defer reopened.Close()
	groups2, err := reopened.GetGroups(257)
	assert.NoError(t, err)
	assert.Len(t, groups2, 1)
	assert.Len(t, groups2[0].ShapeTree.Shape, 2)
	assert.Len(t, groups2[0].ShapeTree.GroupShape, 1)
	assert.Equal(t, "inside nested group", groups2[0].ShapeTree.GroupShape[0].ShapeTree.Shape[0].TextBody.Paragraph[0].Runs[0].Text)
}

// TestGroupHyperLinks checks that the hyperlinks of the shapes of the group
// (and the hyperlink of the group itself) are found by the slide id and the
// shape id, the links can be set and removed for the shapes inside the
// groups (including the nested groups) and the hyperlink relationships of
// the shapes of the groups are not removed from the rels part of the slide
// on save.
func TestGroupHyperLinks(t *testing.T) {
	f, err := OpenFile(groupFixture(t))
	assert.NoError(t, err)
	defer f.Close()

	// the hyperlink of the shape inside the group (the id 101)
	hasLink, link, err := f.GetShapeHyperLink(257, 101)
	assert.NoError(t, err)
	assert.True(t, hasLink)
	assert.Equal(t, groupHyperlinkTarget, link)

	// the hyperlink of the top level group (the id 100)
	hasLink, link, err = f.GetShapeHyperLink(257, 100)
	assert.NoError(t, err)
	assert.True(t, hasLink)
	assert.Equal(t, groupShapeHyperlinkTarget, link)

	// the hyperlink of the shape that is not in the presentation
	_, _, err = f.GetShapeHyperLink(257, 999)
	assert.EqualError(t, err, ErrShapeNotExist{999}.Error())

	// the shapes of the slide with the hyperlinks: the group (the id 100)
	// and the shape inside the group (the id 101)
	ids, err := f.GetShapeHyperLinks(257, "")
	assert.NoError(t, err)
	assert.Equal(t, []int{100, 101}, ids)
	ids, err = f.GetShapeHyperLinks(257, LinkTypeExternal)
	assert.NoError(t, err)
	assert.Equal(t, []int{100, 101}, ids)
	ids, err = f.GetShapeHyperLinks(257, LinkTypeLocation)
	assert.NoError(t, err)
	assert.Empty(t, ids)

	// the hyperlink is set to the shape inside the nested group (the id 104)
	err = f.SetShapeHyperLink(257, 104, "https://pkg.go.dev", LinkTypeExternal)
	assert.NoError(t, err)
	out := filepath.Join(t.TempDir(), "groups.pptx")
	assert.NoError(t, f.SaveAs(out))

	reopened, err := OpenFile(out)
	assert.NoError(t, err)
	defer reopened.Close()
	hasLink, link, err = reopened.GetShapeHyperLink(257, 104)
	assert.NoError(t, err)
	assert.True(t, hasLink)
	assert.Equal(t, "https://pkg.go.dev", link)

	// the hyperlink relationships of the shapes of the groups (the rId2 and
	// the rId3) are not removed from the rels part of the slide on save
	rels := string(readZipPart(t, out, "ppt/slides/_rels/slide2.xml.rels"))
	assert.Contains(t, rels, groupHyperlinkTarget)
	assert.Contains(t, rels, groupShapeHyperlinkTarget)

	// the hyperlink is removed from the shape inside the nested group and
	// the hyperlink relationship is removed from the rels part of the slide
	err = reopened.SetShapeHyperLink(257, 104, "", LinkTypeNone)
	assert.NoError(t, err)
	out2 := filepath.Join(t.TempDir(), "groups2.pptx")
	assert.NoError(t, reopened.SaveAs(out2))

	reopened2, err := OpenFile(out2)
	assert.NoError(t, err)
	defer reopened2.Close()
	hasLink, _, err = reopened2.GetShapeHyperLink(257, 104)
	assert.NoError(t, err)
	assert.False(t, hasLink)
	rels = string(readZipPart(t, out2, "ppt/slides/_rels/slide2.xml.rels"))
	assert.NotContains(t, rels, "https://pkg.go.dev")
	assert.Contains(t, rels, groupHyperlinkTarget)
	assert.Contains(t, rels, groupShapeHyperlinkTarget)

	// the hyperlink of the shape (the group) that is not in the presentation
	err = f.SetShapeHyperLink(257, 999, "https://pkg.go.dev", LinkTypeExternal)
	assert.EqualError(t, err, ErrShapeNotExist{999}.Error())
}

// TestCreateShapeWithGroups checks that the id of the new shape is unique in
// the whole shape tree of the slide: the ids of the shapes of the groups
// (including the nested groups) are taken into account by the next free id.
func TestCreateShapeWithGroups(t *testing.T) {
	f, err := OpenFile(groupFixture(t))
	assert.NoError(t, err)
	defer f.Close()

	// the max id of the shape tree of the slide is the id of the shape
	// inside the group (the id 106)
	shapeID, err := f.CreateShape(257, DecodeShapeProperties{}, DecodeTextBody{})
	assert.NoError(t, err)
	assert.Equal(t, 107, shapeID)

	// the new shape is appended to the end of the shape tree, the group is
	// kept in its place
	shapeOrder, err := f.GetShapeOrder(257)
	assert.NoError(t, err)
	assert.Equal(t, []ShapeOrder{
		{Kind: ShapeOrderSp, ID: 7},
		{Kind: ShapeOrderSp, ID: 8},
		{Kind: ShapeOrderSp, ID: 105},
		{Kind: ShapeOrderGrpSp, ID: 100},
		{Kind: ShapeOrderSp, ID: 106},
		{Kind: ShapeOrderSp, ID: 107},
	}, shapeOrder)

	// the id of the next shape is the max id of the shape tree plus one
	shapeID, err = f.CreateShape(257, DecodeShapeProperties{}, DecodeTextBody{})
	assert.NoError(t, err)
	assert.Equal(t, 108, shapeID)
}
