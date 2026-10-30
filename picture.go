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
	"fmt"
	"image"
	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"
	"math"
	"os"
	"path/filepath"
	"strconv"
	"strings"
)

// GetPictures provides a function to get pictures (p:pic shapes) by given
// slide id, in the document order. The pictures are returned by value:
// edits of the returned slice do not affect the slide.
func (f *File) GetPictures(slideID int) ([]decodePicture, error) {
	slide, err := f.slideReader(slideID)
	if err != nil {
		return nil, err
	}

	return slide.getPictures(), err
}

// getPictures returns the pictures of the slide, in the document order.
func (ds *decodeSlide) getPictures() []decodePicture {
	var pictures []decodePicture
	for i := range ds.CommonSlideData.ShapeTree.Children {
		if child := ds.CommonSlideData.ShapeTree.Children[i].Picture; child != nil {
			pictures = append(pictures, *child)
		}
	}

	return pictures
}

// GetPicture provides a function to get a picture by given slide id and
// picture id (the id attribute of the p:cNvPr element).
func (f *File) GetPicture(slideID int, pictureID int) (decodePicture, error) {
	pictures, err := f.GetPictures(slideID)
	if err != nil {
		return decodePicture{}, err
	}

	for _, pic := range pictures {
		if pic.NonVisualPictureProperties != nil &&
			pic.NonVisualPictureProperties.CommonNonVisualProperties != nil &&
			pic.NonVisualPictureProperties.CommonNonVisualProperties.ID == pictureID {
			return pic, nil
		}
	}

	return decodePicture{}, ErrPictureNotExist{pictureID}
}

// GetImage provides a function to get the image data of the slide referenced
// by the given relationship id (the r:embed value of the picture blip).
func (f *File) GetImage(slideID int, rID string) ([]byte, error) {
	var data []byte
	path, ok := f.getSlideXMLPath(slideID)
	if !ok {
		return data, ErrSlideNotExist{slideID}
	}
	rels, err := f.relsReader(f.getSlideRelsPath(path))
	if err != nil {
		return data, err
	}
	if rels == nil {
		return data, ErrImageNotExist{rID}
	}
	rels.mu.Lock()
	defer rels.mu.Unlock()
	for _, rel := range rels.Relationships {
		if rel.ID == rID && rel.Type == SourceRelationshipImage {
			target := resolvePartPath(path, rel.Target)
			if content, ok := f.Pkg.Load(target); ok {
				return content.([]byte), nil
			}
			return data, ErrImageNotExist{rID}
		}
	}

	return data, ErrImageNotExist{rID}
}

// SetImage provides a function to replace the image data referenced by the
// given relationship id (the r:embed value of the picture blip) in the slide
// with the given bytes. The media part is updated in place, keeping the same
// part path and the original image relationship.
func (f *File) SetImage(slideID int, rID string, data []byte) error {
	if len(data) == 0 {
		return ErrImageEmpty
	}
	path, ok := f.getSlideXMLPath(slideID)
	if !ok {
		return ErrSlideNotExist{slideID}
	}
	rels, err := f.relsReader(f.getSlideRelsPath(path))
	if err != nil {
		return err
	}
	if rels == nil {
		return ErrImageNotExist{rID}
	}
	rels.mu.Lock()
	defer rels.mu.Unlock()
	for _, rel := range rels.Relationships {
		if rel.ID == rID && rel.Type == SourceRelationshipImage {
			f.Pkg.Store(resolvePartPath(path, rel.Target), data)
			return nil
		}
	}

	return ErrImageNotExist{rID}
}

// PictureOptions defines the format settings of the picture on the slide. By
// default (the zero value) the picture is centered on the slide and scaled to
// fit 80% of the slide size, keeping the aspect ratio of the image without
// upscaling.
type PictureOptions struct {
	// Name is the name of the picture shape (the name attribute of the
	// p:cNvPr element).
	Name string
	// OffsetX is the offset of the picture from the left border of the slide
	// in EMU.
	OffsetX int
	// OffsetY is the offset of the picture from the top border of the slide
	// in EMU.
	OffsetY int
	// ScaleX is the scaling ratio of the picture along the X direction
	// (1.0 is the native size of the image at 96 dpi).
	ScaleX float64
	// ScaleY is the scaling ratio of the picture along the Y direction
	// (1.0 is the native size of the image at 96 dpi).
	ScaleY float64
}

// AddImage provides a function to add a new picture to the slide by the given
// file path and the format settings, and returns the id of the picture in the
// slide. The image part is created in the ppt/media folder, the default
// content type for the image extension and the image relationship in the
// slide rels part are updated automatically. For example:
//
//	f, err := gopptx.OpenFile("Book1.pptx")
//	if err != nil {
//	    return err
//	}
//	defer f.Close()
//	// Insert a picture.
//	if _, err = f.AddImage(256, "image.png", nil); err != nil {
//	    return err
//	}
//	// Insert a picture with scaling.
//	if _, err = f.AddImage(256, "image.jpg", &gopptx.PictureOptions{
//	    ScaleX: 0.5,
//	    ScaleY: 0.5,
//	}); err != nil {
//	    return err
//	}
//	return f.SaveAs("Book2.pptx")
func (f *File) AddImage(slideID int, name string, opts *PictureOptions) (int, error) {
	if _, err := os.Stat(name); os.IsNotExist(err) {
		return 0, err
	}
	file, err := os.ReadFile(filepath.Clean(name))
	if err != nil {
		return 0, err
	}
	return f.AddImageFromBytes(slideID, file, opts)
}

// AddImageFromBytes provides a function to add a new picture to the slide by
// the given image data and the format settings and returns the id of the
// picture in the slide. The image part is created in the ppt/media folder,
// the default content type for the image extension and the image
// relationship in the slide rels part are updated automatically. Duplicate
// images are only actually stored once and the pictures that use it will
// reference the same media part.
func (f *File) AddImageFromBytes(slideID int, data []byte, opts *PictureOptions) (int, error) {
	if len(data) == 0 {
		return 0, ErrImageEmpty
	}
	slide, err := f.slideReader(slideID)
	if err != nil {
		return 0, err
	}
	ext, contentType, err := imageTypeFromBytes(data)
	if err != nil {
		return 0, err
	}
	if err = f.ensureImageContentType(ext, contentType); err != nil {
		return 0, err
	}

	partName := f.addImageMedia(data, ext)

	slidePath, _ := f.getSlideXMLPath(slideID)
	// the picture shape references the image relationship (r:embed), so the
	// slide root element must declare the relationships namespace
	f.addNameSpaces(slidePath, SourceRelationship)
	target := "../" + strings.TrimPrefix(partName, "ppt/")
	rID := f.addRels(f.getSlideRelsPath(slidePath), SourceRelationshipImage, target, "")

	pictureID := f.nextShapeID(slide)
	offset, extents := f.pictureLayout(data, opts)
	name := fmt.Sprintf("Picture %d", pictureID)
	if opts != nil && opts.Name != "" {
		name = opts.Name
	}
	preferRelativeResize := 0

	picture := decodePicture{
		NonVisualPictureProperties: &decodeNonVisualPictureProperties{
			CommonNonVisualProperties: &decodeCommonNonVisualProperties{
				ID:   pictureID,
				Name: name,
			},
			CommonNonVisualPictureProperties: &decodeCommonNonVisualPictureProperties{
				PreferRelativeResize: &preferRelativeResize,
			},
			NonVisualProperties: &decodeNonVisualProperties{},
		},
		BlipFill: &decodeBlipFill{
			Blip: &decodeBlip{
				Embed: "rId" + strconv.Itoa(rID),
			},
			Stretch: &decodeStretch{
				FillRect: &FillRect{},
			},
		},
		ShapeProperties: &DecodeShapeProperties{
			Xfrm: &DecodeXfrm{
				Offset:  offset,
				Extents: extents,
			},
			PresetGeometry: &DecodePresetGeometry{
				Preset: "rect",
			},
			NoFill: &noFill{},
			Ln:     &DecodeLine{NoFill: &noFill{}},
		},
	}

	// the new picture is appended to the end of the shape tree: it is
	// drawn above the existing shapes of the slide
	slide.CommonSlideData.ShapeTree.Children = append(slide.CommonSlideData.ShapeTree.Children, decodeShapeTreeChild{Picture: &picture})

	return pictureID, nil
}

// DeleteImage provides a function to delete the picture by the given slide
// id and picture id from the slide. The picture shape (p:pic) is removed from
// the shape tree of the slide, the image relationship is removed from the
// slide rels part when no other shape of the slide uses it, and the media
// part is removed from the package when no other part of the presentation
// references it anymore. The default content type of the image extension is
// removed when no part with this extension is left in the package.
func (f *File) DeleteImage(slideID int, pictureID int) error {
	slide, err := f.slideReader(slideID)
	if err != nil {
		return err
	}
	pictures := slide.getPictures()
	idx := -1
	for i, pic := range pictures {
		if pic.NonVisualPictureProperties != nil &&
			pic.NonVisualPictureProperties.CommonNonVisualProperties != nil &&
			pic.NonVisualPictureProperties.CommonNonVisualProperties.ID == pictureID {
			idx = i
			break
		}
	}
	if idx == -1 {
		return ErrPictureNotExist{pictureID}
	}

	// the picture is taken before the removal
	pic := pictures[idx]
	// remove the picture shape from the shape tree of the slide
	children := slide.CommonSlideData.ShapeTree.Children
	for i := range children {
		if child := children[i].Picture; child != nil &&
			child.NonVisualPictureProperties != nil &&
			child.NonVisualPictureProperties.CommonNonVisualProperties != nil &&
			child.NonVisualPictureProperties.CommonNonVisualProperties.ID == pictureID {
			slide.CommonSlideData.ShapeTree.Children = append(children[:i], children[i+1:]...)
			break
		}
	}

	if pic.BlipFill == nil || pic.BlipFill.Blip == nil || pic.BlipFill.Blip.Embed == "" {
		return nil
	}

	slidePath, _ := f.getSlideXMLPath(slideID)
	// the image relationship is kept while the origin slide XML (outside the
	// p:pic shapes, e.g. the slide background) or the other pictures of the
	// slide still reference it
	embed := pic.BlipFill.Blip.Embed
	used := rawEmbedUse(f.readBytes(slidePath), embed) > 0
	for _, other := range slide.getPictures() {
		if other.BlipFill != nil && other.BlipFill.Blip != nil && other.BlipFill.Blip.Embed == embed {
			used = true
			break
		}
	}
	if used {
		return nil
	}

	rels, err := f.relsReader(f.getSlideRelsPath(slidePath))
	if err != nil {
		return err
	}
	if rels == nil {
		return nil
	}
	rels.mu.Lock()
	defer rels.mu.Unlock()
	for i, rel := range rels.Relationships {
		if rel.ID != embed || rel.Type != SourceRelationshipImage {
			continue
		}
		rels.Relationships = append(rels.Relationships[:i], rels.Relationships[i+1:]...)
		partPath := resolvePartPath(slidePath, rel.Target)
		// remove the media part when no other part of the presentation
		// references it anymore
		if !f.isPartReferenced(partPath) {
			f.Pkg.Delete(partPath)
			f.removeImageContentTypeIfUnused(filepath.Ext(partPath))
		}
		break
	}
	return nil
}

// getSlideRelsPath provides a function to get the path of the slide rels part
// by given slide XML path.
func (f *File) getSlideRelsPath(slidePath string) string {
	return filepath.ToSlash(filepath.Join(filepath.Dir(slidePath), "_rels", filepath.Base(slidePath)+".rels"))
}

// resolvePartPath provides a function to resolve the part path by the given
// base part path and the relationship target (relative or absolute).
func resolvePartPath(basePath, target string) string {
	if strings.HasPrefix(target, "/") {
		return strings.TrimPrefix(target, "/")
	}
	return filepath.ToSlash(filepath.Clean(filepath.Join(filepath.Dir(basePath), target)))
}

// newImagePartName provides a function to generate the next free image part
// name (ppt/media/imageN.ext) in the presentation.
func (f *File) newImagePartName(ext string) string {
	maxIdx := 0
	f.Pkg.Range(func(k, v interface{}) bool {
		name, _ := k.(string)
		if strings.HasPrefix(name, "ppt/media/image") {
			base := strings.TrimPrefix(name, "ppt/media/image")
			if dot := strings.IndexByte(base, '.'); dot > 0 {
				if n, err := strconv.Atoi(base[:dot]); err == nil && n > maxIdx {
					maxIdx = n
				}
			}
		}
		return true
	})
	return fmt.Sprintf("ppt/media/image%d.%s", maxIdx+1, ext)
}

// addImageMedia provides a function to add the image data into the ppt/media
// folder and returns the name of the media part. Duplicate images are only
// actually stored once and the pictures that use it will reference the same
// media part.
func (f *File) addImageMedia(data []byte, ext string) string {
	var name string
	f.Pkg.Range(func(k, v interface{}) bool {
		partName, _ := k.(string)
		if strings.HasPrefix(partName, "ppt/media/") {
			if content, ok := v.([]byte); ok && bytes.Equal(data, content) {
				name = partName
				return false
			}
		}
		return true
	})
	if name != "" {
		return name
	}
	name = f.newImagePartName(ext)
	f.Pkg.Store(name, data)
	return name
}

// isPartReferenced provides a function to check whether the given part is
// still referenced by an image relationship of any part in the presentation.
func (f *File) isPartReferenced(partPath string) bool {
	referenced := false
	seen := make(map[string]bool)
	check := func(name string, v interface{}) {
		var list []relationship
		if r, ok := f.Relationships.Load(name); ok {
			if rels, ok := r.(*relationships); ok {
				list = rels.Relationships
			}
		} else if content, ok := v.([]byte); ok {
			rels := relationships{}
			if err := xml.Unmarshal(namespaceStrictToTransitional(content), &rels); err == nil {
				list = rels.Relationships
			}
		}
		base := basePartOfRelsPart(name)
		for _, rel := range list {
			if rel.Type == SourceRelationshipImage && resolvePartPath(base, rel.Target) == partPath {
				referenced = true
				return
			}
		}
	}
	f.Pkg.Range(func(k, v interface{}) bool {
		name, _ := k.(string)
		if strings.HasSuffix(name, ".rels") {
			seen[name] = true
			check(name, v)
		}
		return !referenced
	})
	if referenced {
		return referenced
	}
	f.Relationships.Range(func(k, v interface{}) bool {
		name, _ := k.(string)
		if strings.HasSuffix(name, ".rels") && !seen[name] {
			check(name, nil)
		}
		return !referenced
	})
	return referenced
}

// basePartOfRelsPart provides a function to get the path of the base part by
// the given rels part path (e.g. ppt/slides/_rels/slide1.xml.rels is mapped
// to ppt/slides/slide1.xml).
func basePartOfRelsPart(relsPath string) string {
	dir := filepath.Dir(relsPath)
	if filepath.Base(dir) != "_rels" {
		return relsPath
	}
	base := strings.TrimSuffix(filepath.Base(relsPath), ".rels")
	return filepath.ToSlash(filepath.Join(filepath.Dir(dir), base))
}

// rawEmbedUse reports the number of occurrences of the given relationship id
// in the raw slide XML outside the p:pic shapes, i.e. in the shapes that the
// library does not model (the slide background, the picture fill of the
// shape).
func rawEmbedUse(raw []byte, embed string) int {
	if len(raw) == 0 || embed == "" {
		return 0
	}
	ref := fmt.Sprintf(`embed="%s"`, embed)
	total := strings.Count(string(raw), ref)
	picUse := 0
	for _, chunk := range bytes.Split(raw, []byte("<p:pic"))[1:] {
		if strings.Contains(string(chunk), ref) {
			picUse++
		}
	}
	return total - picUse
}

// removeImageContentTypeIfUnused provides a function to remove the default
// content type of the given extension from the content types part when no
// part with this extension is left in the package. The defaults declared in
// the origin content types part are kept.
func (f *File) removeImageContentTypeIfUnused(ext string) {
	if ext == "" {
		return
	}
	ext = strings.TrimPrefix(ext, ".")
	inUse := false
	f.Pkg.Range(func(k, v interface{}) bool {
		name, _ := k.(string)
		if strings.EqualFold(filepath.Ext(name), "."+ext) {
			inUse = true
			return false
		}
		return true
	})
	if inUse {
		return
	}
	content, err := f.contentTypesReader()
	if err != nil {
		return
	}
	content.mu.Lock()
	defer content.mu.Unlock()
	for i, d := range content.Defaults {
		if !strings.EqualFold(d.Extension, ext) {
			continue
		}
		// keep the default declared in the origin content types part
		if inStrSlice(f.originalContentTypeExtensions, d.Extension, false) != -1 {
			return
		}
		content.Defaults = append(content.Defaults[:i], content.Defaults[i+1:]...)
		return
	}
}

// ensureImageContentType provides a function to ensure that the content types
// part contains a default content type for the given image extension.
func (f *File) ensureImageContentType(ext, contentType string) error {
	content, err := f.contentTypesReader()
	if err != nil {
		return err
	}
	content.mu.Lock()
	defer content.mu.Unlock()
	for i, d := range content.Defaults {
		if strings.EqualFold(d.Extension, ext) {
			content.Defaults[i].ContentType = contentType
			return nil
		}
	}
	content.Defaults = append(content.Defaults, contentTypeDefault{
		Extension:   ext,
		ContentType: contentType,
	})
	return nil
}

// nextShapeID provides a function to get the next free shape id in the slide
// shape tree, taking into account all the shape types (sp, pic, cxnSp and
// graphicFrame).
func (f *File) nextShapeID(slide *decodeSlide) int {
	id := defaultXMLShapeID
	for _, s := range slide.getShapes() {
		if s.NonVisualShapeProperties != nil && s.NonVisualShapeProperties.CommonNonVisualProperties != nil {
			id = max(id, s.NonVisualShapeProperties.CommonNonVisualProperties.ID)
		}
	}

	for _, pic := range slide.getPictures() {
		if pic.NonVisualPictureProperties != nil && pic.NonVisualPictureProperties.CommonNonVisualProperties != nil {
			id = max(id, pic.NonVisualPictureProperties.CommonNonVisualProperties.ID)
		}
	}

	for _, conn := range slide.getConnectors() {
		if conn.NonVisualConnectorProperties != nil && conn.NonVisualConnectorProperties.CommonNonVisualProperties != nil {
			id = max(id, conn.NonVisualConnectorProperties.CommonNonVisualProperties.ID)
		}
	}

	for _, gf := range slide.getGraphicFrames() {
		if gf.NonVisualGraphicFrameProperties != nil && gf.NonVisualGraphicFrameProperties.CommonNonVisualProperties != nil {
			id = max(id, gf.NonVisualGraphicFrameProperties.CommonNonVisualProperties.ID)
		}
	}

	return id + 1
}

// pictureLayout provides a function to calculate the offset and extents (in
// EMU) of a new picture. By default the image is scaled to fit 80% of the
// slide size (keeping the aspect ratio, without upscaling) and centered on
// the slide. The position and the scale of the picture can be overridden by
// the given picture options.
func (f *File) pictureLayout(data []byte, opts *PictureOptions) (*Offset, *Extents) {
	slideSize := slideSize{CY: 5143500, CX: 9144000}
	if presentation, _ := f.presentationReader(); presentation != nil && presentation.SlideSize != nil {
		slideSize = *presentation.SlideSize
	}

	const emuPerPx = 9525 // EMU per pixel at 96 dpi
	cx, cy := slideSize.CX/2, slideSize.CY/2
	if cfg, _, err := image.DecodeConfig(bytes.NewReader(data)); err == nil && cfg.Width > 0 && cfg.Height > 0 {
		cx, cy = cfg.Width*emuPerPx, cfg.Height*emuPerPx
	} else if w, h, ok := svgSize(data); ok && w > 0 && h > 0 {
		// the SVG image size is taken from the width and height attributes
		// (or the viewBox) of the root svg element
		cx, cy = w*emuPerPx, h*emuPerPx
	}
	if opts != nil && (opts.ScaleX > 0 || opts.ScaleY > 0) {
		scaleX, scaleY := opts.ScaleX, opts.ScaleY
		if scaleX == 0 {
			scaleX = scaleY
		}
		if scaleY == 0 {
			scaleY = scaleX
		}
		cx, cy = int(float64(cx)*scaleX), int(float64(cy)*scaleY)
	} else {
		const fitFraction = 0.8 // the picture fits 80% of the slide size
		maxCX := int(float64(slideSize.CX) * fitFraction)
		maxCY := int(float64(slideSize.CY) * fitFraction)
		if float64(cx) > float64(maxCX) || float64(cy) > float64(maxCY) {
			scale := math.Min(float64(maxCX)/float64(cx), float64(maxCY)/float64(cy))
			cx, cy = int(float64(cx)*scale), int(float64(cy)*scale)
		}
	}

	offset := &Offset{
		X: (slideSize.CX - cx) / 2,
		Y: (slideSize.CY - cy) / 2,
	}
	if opts != nil && (opts.OffsetX != 0 || opts.OffsetY != 0) {
		offset = &Offset{X: opts.OffsetX, Y: opts.OffsetY}
	}
	return offset, &Extents{
		CX: cx,
		CY: cy,
	}
}

// imageTypeFromBytes provides a function to detect the image format by the
// given bytes. It returns the file extension (without the dot) and the content
// type of the image.
func imageTypeFromBytes(data []byte) (ext string, contentType string, err error) {
	switch {
	case isJPEG(data):
		return "jpg", "image/jpeg", nil
	case isPNG(data):
		return "png", "image/png", nil
	case isGIF(data):
		return "gif", "image/gif", nil
	case isBMP(data):
		return "bmp", "image/bmp", nil
	case isTIFF(data):
		return "tiff", "image/tiff", nil
	case isWEBP(data):
		return "webp", "image/webp", nil
	case isICO(data):
		return "ico", "image/x-icon", nil
	case isEMF(data):
		return "emf", "image/x-emf", nil
	case isWMF(data):
		return "wmf", "image/x-wmf", nil
	case isSVG(data):
		return "svg", "image/svg+xml", nil
	}
	return "", "", ErrUnsupportedImageFormat
}

func isJPEG(data []byte) bool {
	return len(data) >= 3 && data[0] == 0xFF && data[1] == 0xD8 && data[2] == 0xFF
}

func isPNG(data []byte) bool {
	return len(data) >= 8 && bytes.Equal(data[:8], []byte{0x89, 'P', 'N', 'G', 0x0D, 0x0A, 0x1A, 0x0A})
}

func isGIF(data []byte) bool {
	return len(data) >= 6 && (bytes.Equal(data[:6], []byte("GIF87a")) || bytes.Equal(data[:6], []byte("GIF89a")))
}

func isBMP(data []byte) bool {
	return len(data) >= 2 && data[0] == 'B' && data[1] == 'M'
}

func isTIFF(data []byte) bool {
	return len(data) >= 4 && (bytes.Equal(data[:4], []byte{'I', 'I', '*', 0x00}) || bytes.Equal(data[:4], []byte{'M', 'M', 0x00, '*'}))
}

func isWEBP(data []byte) bool {
	return len(data) >= 12 && bytes.Equal(data[:4], []byte("RIFF")) && bytes.Equal(data[8:12], []byte("WEBP"))
}

func isICO(data []byte) bool {
	return len(data) >= 4 && data[0] == 0x00 && data[1] == 0x00 && data[2] == 0x01 && data[3] == 0x00
}

func isEMF(data []byte) bool {
	return len(data) >= 4 && data[0] == 0x01 && data[1] == 0x00 && data[2] == 0x00 && data[3] == 0x00
}

func isWMF(data []byte) bool {
	return len(data) >= 4 && (bytes.Equal(data[:4], []byte{0xD7, 0xCD, 0xC6, 0x9A}) || bytes.Equal(data[:4], []byte{0x01, 0x00, 0x09, 0x00}))
}

func isSVG(data []byte) bool {
	if !bytes.HasPrefix(data, []byte("<?xml")) && !bytes.HasPrefix(data, []byte("<svg")) {
		return false
	}
	limit := len(data)
	if limit > 1024 {
		limit = 1024
	}
	return bytes.Contains(data[:limit], []byte("svg"))
}

// svgSize provides a function to get the size of an SVG image in pixels by
// the width and height attributes of the root svg element. The viewBox
// attribute is used when one of them is absent.
func svgSize(data []byte) (width, height int, ok bool) {
	start := bytes.Index(data, []byte("<svg"))
	if start == -1 {
		return 0, 0, false
	}
	end := bytes.IndexByte(data[start:], '>')
	if end == -1 {
		return 0, 0, false
	}
	tag := data[start : start+end]
	if w, wOK := svgAttrPixels(tag, "width"); wOK {
		if h, hOK := svgAttrPixels(tag, "height"); hOK {
			return w, h, true
		}
	}
	// the viewBox attribute: "min-x min-y width height"
	if i := bytes.Index(tag, []byte("viewBox=")); i != -1 {
		rest := tag[i+len("viewBox="):]
		j := bytes.IndexAny(rest, "\"'")
		if j == -1 {
			return 0, 0, false
		}
		fields := strings.Fields(string(rest[:j]))
		if len(fields) == 4 {
			w, err1 := strconv.ParseFloat(fields[2], 64)
			h, err2 := strconv.ParseFloat(fields[3], 64)
			if err1 == nil && err2 == nil && w > 0 && h > 0 {
				return int(w + 0.5), int(h + 0.5), true
			}
		}
	}
	return 0, 0, false
}

// svgAttrPixels provides a function to get the pixel value of the given
// attribute of the root svg tag. The in, cm, mm, pt and pc units are
// converted to the pixels at 96 dpi.
func svgAttrPixels(tag []byte, name string) (int, bool) {
	for _, quote := range []string{`"`, "'"} {
		marker := []byte(name + "=" + quote)
		i := bytes.Index(tag, marker)
		if i == 0 || (tag[i-1] != ' ' && tag[i-1] != '\t') {
			continue
		}
		rest := tag[i+len(marker):]
		j := bytes.IndexByte(rest, quote[0])
		if j == -1 {
			return 0, false
		}
		value := string(rest[:j])
		k := 0
		for k < len(value) {
			c := value[k]
			if (c >= '0' && c <= '9') || c == '.' || c == '-' || c == '+' {
				k++
				continue
			}
			break
		}
		num, unit := value[:k], strings.ToLower(value[k:])
		f, err := strconv.ParseFloat(num, 64)
		if err != nil || f <= 0 {
			return 0, false
		}
		switch unit {
		case "", "px":
		case "in":
			f *= 96
		case "cm":
			f *= 96 / 2.54
		case "mm":
			f *= 96 / 25.4
		case "pt":
			f *= 96 / 72
		case "pc":
			f *= 96 / 6
		default:
			return 0, false
		}
		return int(f + 0.5), true
	}
	return 0, false
}
