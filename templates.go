// Copyright 2026 kenny-not-dead. All rights reserved.
// Use of this source code is governed by a BSD-style license
// that can be found in the LICENSE file.
//
// Package gopptx provides functionality to create and manipulate PowerPoint
// (.pptx) files in Go, using the Office Open XML (ECMA-376) format.

package gopptx

import (
	_ "embed"
)

var (
	//go:embed templates/\[Content_Types].xml
	templateContentTypes string

	//go:embed templates/\app.xml
	templateDocpropsApp string

	//go:embed templates/core.xml
	templateDocpropsCore string

	//go:embed templates/slide1.xml
	TemplateSlide string

	//go:embed templates/slide1-rels.xml
	TemplateSlideRels string

	//go:embed templates/slideLayout1.xml
	templateSlideLayout string

	//go:embed templates/slideLayout1-rels.xml
	templateSlideLayoutRels string

	//go:embed templates/slideMaster1.xml
	templateSlideMaster string

	//go:embed templates/slideMaster1-rels.xml
	templateSlideMasterRels string

	//go:embed templates/presentation.xml
	templatePresentation string

	//go:embed templates/presProps.xml
	templatePresProps string

	//go:embed templates/presentation-rels.xml
	templatePresentationRels string

	//go:embed templates/rels.xml
	templateRels string

	//go:embed templates/theme1.xml
	templateTheme string
)
