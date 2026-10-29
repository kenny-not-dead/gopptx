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
	"io"
	"path/filepath"
	"strconv"
	"strings"
)

// NewNotes provides a function to create a new notes slide for the given
// slide. The notes slide part (ppt/notesSlides/notesSlideN.xml) is created
// with the standard placeholders (the slide image, the notes body and the
// slide number), the notesSlide relationship is added to the slide rels part
// and the content types are updated. The notes master and the second theme
// are created automatically when the presentation has no notes master yet.
func (f *File) NewNotes(slideID int) error {
	if _, ok := f.getSlideXMLPath(slideID); !ok {
		return ErrSlideNotExist{slideID}
	}

	if _, ok := f.getNotesSlidePath(slideID); ok {
		return ErrNotesSlideExists{slideID}
	}

	if err := f.ensureNotesMaster(); err != nil {
		return err
	}

	slidePath, _ := f.getSlideXMLPath(slideID)
	notesPath := f.newNotesSlidePath()
	notesBase := filepath.Base(notesPath)
	notesRelsPath := "ppt/notesSlides/_rels/" + notesBase + ".rels"
	slideBase := filepath.Base(slidePath)

	// create the notes slide part and its rels part (the notes slide
	// references the slide and the notes master)
	f.Pkg.Store(notesPath, []byte(xml.Header+TemplateNotesSlide))
	// capture the origin root element attributes of the notes slide
	// template, so the new notes slide keeps the template namespaces on save
	f.xmlAttr.Store(notesPath, getRootElement(f.xmlNewDecoder(strings.NewReader(TemplateNotesSlide))))
	notes := new(decodeNotesSlide)
	if err := xml.Unmarshal([]byte(TemplateNotesSlide), notes); err != nil {
		return err
	}

	f.NotesSlides.Store(notesPath, notes)
	notesRels := &relationships{
		Relationships: []relationship{
			{
				ID:     "rId1",
				Type:   SourceRelationshipNotesMaster,
				Target: f.notesMasterRelsTarget(),
			},
			{
				ID:     "rId2",
				Type:   SourceRelationshipSlide,
				Target: "../slides/" + slideBase,
			},
		},
	}

	notesRelsOutput, _ := xml.Marshal(notesRels)
	f.Pkg.Store(notesRelsPath, append([]byte(xml.Header), selfCloseEmptyElements(notesRelsOutput)...))

	// add the notesSlide relationship to the slide rels part
	f.addRels(f.getSlideRelsPath(slidePath), SourceRelationshipNotesSlide, "../notesSlides/"+notesBase, "")

	// Update [Content_Types].xml
	_ = f.setContentTypes("/ppt/notesSlides/_rels/"+notesBase+".rels", ContentTypeRelationships)
	_ = f.setContentTypes("/ppt/notesSlides/"+notesBase, ContentTypeNotesSlideML)

	return nil
}

// GetNotes provides a function to get the notes (the text of the notes body
// placeholder) by the given slide id.
func (f *File) GetNotes(slideID int) (DecodeTextBody, error) {
	notes, err := f.notesSlideReader(slideID)
	if err != nil {
		return DecodeTextBody{}, err
	}

	shape := notes.notesBodyShape()
	if shape == nil || shape.TextBody == nil {
		return DecodeTextBody{}, nil
	}

	return *shape.TextBody, nil
}

// SetNotes provides a function to set the notes (the text of the notes body
// placeholder) by the given slide id. The notes slide is created
// automatically when the slide has no notes yet.
func (f *File) SetNotes(slideID int, textBody DecodeTextBody) error {
	notes, err := f.notesSlideReader(slideID)
	if err != nil {
		var notExist ErrNotesSlideNotExist
		if !errors.As(err, &notExist) {
			return err
		}
		if err := f.NewNotes(slideID); err != nil {
			return err
		}
		notes, err = f.notesSlideReader(slideID)
		if err != nil {
			return err
		}
	}

	shape := notes.notesBodyShape()
	if shape == nil {
		return ErrNotesBodyPlaceholderNotExist
	}

	shape.TextBody = &textBody

	return nil
}

// DeleteNotes provides a function to delete the notes slide of the given
// slide: the notes slide part and its rels part are removed from the package,
// the notesSlide relationship is removed from the slide rels part and the
// content types are updated.
func (f *File) DeleteNotes(slideID int) error {
	notesPath, ok := f.getNotesSlidePath(slideID)
	if !ok {
		return ErrNotesSlideNotExist{slideID}
	}
	slidePath, _ := f.getSlideXMLPath(slideID)
	notesBase := filepath.Base(notesPath)
	notesRelsPath := "ppt/notesSlides/_rels/" + notesBase + ".rels"

	// remove the notesSlide relationship from the slide rels part
	rels, _ := f.relsReader(f.getSlideRelsPath(slidePath))
	if rels != nil {
		rels.mu.Lock()
		for idx, rel := range rels.Relationships {
			if rel.Type == SourceRelationshipNotesSlide &&
				resolvePartPath(slidePath, rel.Target) == notesPath {
				rels.Relationships = append(rels.Relationships[:idx], rels.Relationships[idx+1:]...)
				break
			}
		}
		rels.mu.Unlock()
	}

	// Update [Content_Types].xml
	_ = f.removeContentTypesPart(ContentTypeNotesSlideML, notesPath)
	_ = f.removeContentTypesPart(ContentTypeRelationships, "ppt/notesSlides/_rels/"+notesBase+".rels")

	f.Pkg.Delete(notesPath)
	f.Pkg.Delete(notesRelsPath)
	f.Relationships.Delete(notesRelsPath)
	f.NotesSlides.Delete(notesPath)
	f.xmlAttr.Delete(notesPath)

	return nil
}

// getNotesSlidePath provides a function to get the notes slide part path by
// the given slide id from the notesSlide relationship in the slide rels part.
func (f *File) getNotesSlidePath(slideID int) (string, bool) {
	slidePath, ok := f.getSlideXMLPath(slideID)
	if !ok {
		return "", false
	}

	rels, _ := f.relsReader(f.getSlideRelsPath(slidePath))
	if rels == nil {
		return "", false
	}

	rels.mu.Lock()
	defer rels.mu.Unlock()
	for _, rel := range rels.Relationships {
		if rel.Type == SourceRelationshipNotesSlide {
			return resolvePartPath(slidePath, rel.Target), true
		}
	}

	return "", false
}

// notesSlideReader provides a function to get the parsed notes slide by the
// given slide id. The notes slide is loaded lazily and is not re-serialized
// on save when it is not modified.
func (f *File) notesSlideReader(slideID int) (*decodeNotesSlide, error) {
	path, ok := f.getNotesSlidePath(slideID)
	if !ok {
		return nil, ErrNotesSlideNotExist{slideID}
	}
	if notes, ok := f.NotesSlides.Load(path); ok && notes != nil {
		return notes.(*decodeNotesSlide), nil
	}

	notes := new(decodeNotesSlide)
	if _, ok := f.xmlAttr.Load(path); !ok {
		d := f.xmlNewDecoder(bytes.NewReader(f.readBytes(path)))
		f.xmlAttr.Store(path, getRootElement(d))
	}

	if err := f.xmlNewDecoder(bytes.NewReader(f.readBytes(path))).
		Decode(notes); err != nil && err != io.EOF {
		return nil, err
	}

	if _, ok := f.checked.Load(path); !ok {
		f.checked.Store(path, true)
	}

	f.NotesSlides.Store(path, notes)

	return notes, nil
}

// ensureNotesMaster provides a function to create the notes master part and
// the second theme in the package when the presentation has no notes master
// yet, and to register the notes master in the presentation.xml.
func (f *File) ensureNotesMaster() error {
	presentation, err := f.presentationReader()
	if err != nil {
		return err
	}

	if presentation.NotesMaster.NotesMaster.RelationshipID != "" {
		return nil
	}

	// the presentation rels may already reference a notes master part:
	// reuse the part and its relationship
	rels, _ := f.relsReader(f.getPresentationRelsPath())
	notesMasterRID := ""
	if rels != nil {
		rels.mu.Lock()
		for _, rel := range rels.Relationships {
			if rel.Type == SourceRelationshipNotesMaster {
				notesMasterRID = rel.ID
				break
			}
		}
		rels.mu.Unlock()
	}

	if notesMasterRID == "" {
		// create the notes master part, its rels part and the second theme
		f.Pkg.Store(defaultXMLPathNotesMaster, []byte(xml.Header+templateNotesMaster))
		f.Pkg.Store(defaultXMLPathNotesMasterRels, []byte(xml.Header+templateNotesMasterRels))
		f.Pkg.Store(defaultXMLPathTheme2, []byte(xml.Header+templateTheme))
		// Update [Content_Types].xml
		_ = f.setContentTypes("/ppt/notesMasters/notesMaster1.xml", ContentTypeNotesMasterML)
		_ = f.setContentTypes("/ppt/notesMasters/_rels/notesMaster1.xml.rels", ContentTypeRelationships)
		_ = f.setContentTypes("/ppt/theme/theme2.xml", ContentTypeThemeML)
		// add the notesMaster relationship to the presentation rels part
		notesMasterRID = "rId" + strconv.Itoa(f.addRels(f.getPresentationRelsPath(), SourceRelationshipNotesMaster, "notesMasters/notesMaster1.xml", ""))
	}
	presentation.NotesMaster = decodeNotesMasterList{
		NotesMaster: decodeNotesMasterID{RelationshipID: notesMasterRID},
	}

	return nil
}

// notesMasterRelsTarget provides a function to get the target of the notes
// master part relative to the notes slides folder (../notesMasters/...).
func (f *File) notesMasterRelsTarget() string {
	rels, _ := f.relsReader(f.getPresentationRelsPath())
	if rels != nil {
		rels.mu.Lock()
		for _, rel := range rels.Relationships {
			if rel.Type == SourceRelationshipNotesMaster {
				target := resolvePartPath(f.getPresentationPath(), rel.Target)
				rels.mu.Unlock()
				return "../" + strings.TrimPrefix(target, "ppt/")
			}
		}
		rels.mu.Unlock()
	}

	return "../" + strings.TrimPrefix(defaultXMLPathNotesMaster, "ppt/")
}

// newNotesSlidePath provides a function to generate the next free notes slide
// part name (ppt/notesSlides/notesSlideN.xml) in the presentation.
func (f *File) newNotesSlidePath() string {
	maxIdx := 0
	findMax := func(name string) {
		if strings.HasPrefix(name, "ppt/notesSlides/notesSlide") && strings.HasSuffix(name, ".xml") {
			base := strings.TrimSuffix(strings.TrimPrefix(name, "ppt/notesSlides/notesSlide"), ".xml")
			if n, err := strconv.Atoi(base); err == nil && n > maxIdx {
				maxIdx = n
			}
		}
	}

	f.Pkg.Range(func(k, v interface{}) bool {
		name, _ := k.(string)
		findMax(name)
		return true
	})

	f.tempFiles.Range(func(k, v interface{}) bool {
		name, _ := k.(string)
		findMax(name)
		return true
	})

	return "ppt/notesSlides/notesSlide" + strconv.Itoa(maxIdx+1) + ".xml"
}
