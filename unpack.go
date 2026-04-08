/*
   Copyright (c) 2020 gingfrederik
   Copyright (c) 2021 Gonzalo Fernandez-Victorio
   Copyright (c) 2021 Basement Crowd Ltd (https://www.basementcrowd.com)
   Copyright (c) 2023 Fumiama Minamoto (源文雨)

   This program is free software: you can redistribute it and/or modify
   it under the terms of the GNU Affero General Public License as published
   by the Free Software Foundation, either version 3 of the License, or
   (at your option) any later version.

   This program is distributed in the hope that it will be useful,
   but WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
   GNU Affero General Public License for more details.

   You should have received a copy of the GNU Affero General Public License
   along with this program.  If not, see <https://www.gnu.org/licenses/>.
*/

package docx

import (
	"archive/zip"
	"encoding/xml"
	"errors"
	"io"
	"strconv"
	"strings"
)

// unpack receives a zip file (word documents are a zip with multiple xml inside)
// and parses the files that are relevant for us:
//
//  1. Document
//  2. Relationships
//  3. Media
//
// Then it stores all other files into tmpfslist for packing.
func unpack(zipReader *zip.Reader) (docx *Docx, err error) {
	docx = new(Docx)
	docx.mediaNameIdx = make(map[string]int, 64)
	docx.slowIDs = make(map[string]uintptr, 64)
	docx.headers = make(map[HeaderKind]*Header, 3)
	docx.footers = make(map[FooterKind]*Footer, 3)
	docx.tmplfs = zipReader
	docx.tmpfslst = make([]string, 0, 64)
	headerFiles := make([]*zip.File, 0, 4)
	footerFiles := make([]*zip.File, 0, 4)
	for _, f := range zipReader.File {
		if f.Name == "word/_rels/document.xml.rels" {
			err = docx.parseDocRelation(f)
			if err != nil {
				return
			}
			continue
		}
		if f.Name == "word/document.xml" {
			err = docx.parseDocument(f)
			if err != nil {
				return
			}
			continue
		}
		if strings.HasPrefix(f.Name, MEDIA_FOLDER) {
			err = docx.parseMedia(f)
			if err != nil {
				return
			}
			continue
		}
		if strings.HasPrefix(f.Name, "word/header") && strings.HasSuffix(strings.ToLower(f.Name), ".xml") {
			headerFiles = append(headerFiles, f)
			docx.tmpfslst = append(docx.tmpfslst, f.Name)
			continue
		}
		if strings.HasPrefix(f.Name, "word/footer") && strings.HasSuffix(strings.ToLower(f.Name), ".xml") {
			footerFiles = append(footerFiles, f)
			docx.tmpfslst = append(docx.tmpfslst, f.Name)
			continue
		}
		// fill remaining files into tmpfslst
		docx.tmpfslst = append(docx.tmpfslst, f.Name)
	}
	if err = docx.parseHeaderFooterParts(headerFiles, footerFiles); err != nil {
		return nil, err
	}
	//TODO: find last imageID
	docx.imageID = 100000
	return
}

// parseDocument processes one of the relevant files, the one with the actual document
func (f *Docx) parseDocument(file *zip.File) error {
	zf, err := file.Open()
	if err != nil {
		return err
	}
	defer zf.Close()

	f.Document.XMLW = XMLNS_W
	f.Document.XMLR = XMLNS_R
	f.Document.XMLWP = XMLNS_WP
	// f.Document.XMLMC = XMLNS_MC
	// f.Document.XMLO = XMLNS_O
	// f.Document.XMLV = XMLNS_V
	f.Document.XMLWPS = XMLNS_WPS
	f.Document.XMLWPC = XMLNS_WPC
	f.Document.XMLWPG = XMLNS_WPG
	// f.Document.XMLWP14 = XMLNS_WP14
	f.Document.XMLName.Space = XMLNS_W
	f.Document.XMLName.Local = "document"

	f.Document.Body.file = f
	//TODO: find last docID
	f.docID = 100000
	err = xml.NewDecoder(zf).Decode(&f.Document)
	return err
}

// parseDocRelation processes one of the relevant files, the one with the relationships
func (f *Docx) parseDocRelation(file *zip.File) error {
	zf, err := file.Open()
	if err != nil {
		return err
	}
	defer zf.Close()

	f.docRelation.Xmlns = XMLNS_R
	err = xml.NewDecoder(zf).Decode(&f.docRelation)
	if err != nil {
		return err
	}
	for _, r := range f.docRelation.Relationship {
		if !strings.HasPrefix(r.ID, "rId") {
			return errors.New("invalid rel ID: " + r.ID)
		}
		id, err := strconv.ParseUint(r.ID[3:], 10, 64)
		if err != nil {
			return err
		}
		if f.rID < uintptr(id) {
			f.rID = uintptr(id)
		}
	}
	return nil
}

// parseMedia add the media into Docx struct
func (f *Docx) parseMedia(file *zip.File) error {
	name := file.Name[len(MEDIA_FOLDER):]
	zf, err := file.Open()
	if err != nil {
		return err
	}
	data, err := io.ReadAll(zf)
	if err != nil {
		return err
	}
	f.mediaNameIdx[name] = len(f.media)
	f.media = append(f.media, Media{Name: name, Data: data})
	return zf.Close()
}

func (f *Docx) parseHeaderFooterParts(headerFiles, footerFiles []*zip.File) error {
	parsedHeaders := make(map[string]*Header, len(headerFiles))
	for _, zf := range headerFiles {
		h, err := f.parseHeaderFile(zf)
		if err != nil {
			return err
		}
		parsedHeaders[strings.TrimPrefix(zf.Name, "word/")] = h
	}
	parsedFooters := make(map[string]*Footer, len(footerFiles))
	for _, zf := range footerFiles {
		ft, err := f.parseFooterFile(zf)
		if err != nil {
			return err
		}
		parsedFooters[strings.TrimPrefix(zf.Name, "word/")] = ft
	}

	mainSect := f.ensureMainSectPr(false)
	if mainSect == nil {
		return nil
	}
	for _, ref := range mainSect.HeaderRefs {
		if ref == nil || ref.RID == "" {
			continue
		}
		rel := f.findRelationshipByID(ref.RID)
		if rel == nil {
			continue
		}
		h := parsedHeaders[normalizeRelTarget(rel.Target)]
		if h == nil {
			continue
		}
		f.headers[normalizeHeaderKind(HeaderKind(ref.Type))] = h
	}
	for _, ref := range mainSect.FooterRefs {
		if ref == nil || ref.RID == "" {
			continue
		}
		rel := f.findRelationshipByID(ref.RID)
		if rel == nil {
			continue
		}
		ft := parsedFooters[normalizeRelTarget(rel.Target)]
		if ft == nil {
			continue
		}
		f.footers[normalizeFooterKind(FooterKind(ref.Type))] = ft
	}
	return nil
}

func (f *Docx) parseHeaderFile(file *zip.File) (*Header, error) {
	rc, err := file.Open()
	if err != nil {
		return nil, err
	}
	defer rc.Close()
	var h Header
	h.file = f
	if err := xml.NewDecoder(rc).Decode(&h); err != nil {
		return nil, err
	}
	return &h, nil
}

func (f *Docx) parseFooterFile(file *zip.File) (*Footer, error) {
	rc, err := file.Open()
	if err != nil {
		return nil, err
	}
	defer rc.Close()
	var ft Footer
	ft.file = f
	if err := xml.NewDecoder(rc).Decode(&ft); err != nil {
		return nil, err
	}
	return &ft, nil
}
