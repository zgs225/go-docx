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
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"strings"
)

// pack receives a zip file writer (word documents are a zip with multiple xml inside)
// and writes the relevant files. Some of them come from the empty_constants file,
// others from the actual in-memory structure
func (f *Docx) pack(zipWriter *zip.Writer) (err error) {
	files := make(map[string]io.Reader, 64)

	if f.template != "" {
		for _, name := range f.tmpfslst {
			files[name], err = f.tmplfs.Open("xml/" + f.template + "/" + name)
			if err != nil {
				return
			}
		}
	} else {
		for _, name := range f.tmpfslst {
			files[name], err = f.tmplfs.Open(name)
			if err != nil {
				return
			}
		}
	}

	headerPaths, footerPaths, err := f.syncHeaderFooterForPack(files)
	if err != nil {
		return err
	}
	if len(headerPaths) > 0 || len(footerPaths) > 0 {
		if err := ensureContentTypesHeaderFooterOverrides(files, headerPaths, footerPaths); err != nil {
			return err
		}
	}

	files["word/_rels/document.xml.rels"] = marshaller{data: &f.docRelation}
	files["word/document.xml"] = marshaller{data: &f.Document}

	for _, m := range f.media {
		files[m.String()] = bytes.NewReader(m.Data)
	}

	for path, r := range files {
		w, err := zipWriter.Create(path)
		if err != nil {
			return err
		}

		_, err = io.Copy(w, r)
		if err != nil {
			return err
		}
	}

	return
}

func (f *Docx) syncHeaderFooterForPack(files map[string]io.Reader) ([]string, []string, error) {
	if len(f.headers) == 0 && len(f.footers) == 0 {
		return nil, nil, nil
	}
	mainSect := f.ensureMainSectPr(true)
	existingHeaderRID := make(map[HeaderKind]string, len(mainSect.HeaderRefs))
	for _, ref := range mainSect.HeaderRefs {
		if ref == nil {
			continue
		}
		existingHeaderRID[normalizeHeaderKind(HeaderKind(ref.Type))] = ref.RID
	}
	existingFooterRID := make(map[FooterKind]string, len(mainSect.FooterRefs))
	for _, ref := range mainSect.FooterRefs {
		if ref == nil {
			continue
		}
		existingFooterRID[normalizeFooterKind(FooterKind(ref.Type))] = ref.RID
	}

	headerPaths := make([]string, 0, len(f.headers))
	footerPaths := make([]string, 0, len(f.footers))
	hrefs := make([]*HeaderReference, 0, len(f.headers))
	frefs := make([]*FooterReference, 0, len(f.footers))

	for _, kind := range headerKindsInOrder() {
		h := f.headers[kind]
		if h == nil {
			continue
		}
		h.file = f
		path := fmt.Sprintf("word/header_%s.xml", kind)
		if rid := existingHeaderRID[kind]; rid != "" {
			if rel := f.findRelationshipByID(rid); rel != nil {
				path = "word/" + normalizeRelTarget(rel.Target)
			}
		}
		files[path] = marshaller{data: h}
		target := strings.TrimPrefix(path, "word/")
		rid := f.ensureInternalPartRelation(REL_HEADER, target, existingHeaderRID[kind])
		hrefs = append(hrefs, &HeaderReference{Type: string(kind), RID: rid})
		headerPaths = append(headerPaths, path)
	}
	for _, kind := range footerKindsInOrder() {
		ft := f.footers[kind]
		if ft == nil {
			continue
		}
		ft.file = f
		path := fmt.Sprintf("word/footer_%s.xml", kind)
		if rid := existingFooterRID[kind]; rid != "" {
			if rel := f.findRelationshipByID(rid); rel != nil {
				path = "word/" + normalizeRelTarget(rel.Target)
			}
		}
		files[path] = marshaller{data: ft}
		target := strings.TrimPrefix(path, "word/")
		rid := f.ensureInternalPartRelation(REL_FOOTER, target, existingFooterRID[kind])
		frefs = append(frefs, &FooterReference{Type: string(kind), RID: rid})
		footerPaths = append(footerPaths, path)
	}

	mainSect.setHeaderFooterRefs(hrefs, frefs)
	return headerPaths, footerPaths, nil
}

func ensureContentTypesHeaderFooterOverrides(files map[string]io.Reader, headerPaths, footerPaths []string) error {
	const contentTypesPath = "[Content_Types].xml"
	r, ok := files[contentTypesPath]
	if !ok {
		return nil
	}
	data, err := io.ReadAll(r)
	if err != nil {
		return err
	}
	updated := string(data)
	insert := make([]string, 0, len(headerPaths)+len(footerPaths))
	for _, p := range headerPaths {
		part := "/" + strings.TrimPrefix(p, "/")
		if !strings.Contains(updated, `PartName="`+part+`"`) {
			insert = append(insert, `<Override PartName="`+part+`" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>`)
		}
	}
	for _, p := range footerPaths {
		part := "/" + strings.TrimPrefix(p, "/")
		if !strings.Contains(updated, `PartName="`+part+`"`) {
			insert = append(insert, `<Override PartName="`+part+`" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>`)
		}
	}
	if len(insert) == 0 {
		files[contentTypesPath] = bytes.NewReader(data)
		return nil
	}
	anchor := "</Types>"
	idx := strings.LastIndex(updated, anchor)
	if idx < 0 {
		files[contentTypesPath] = bytes.NewReader(data)
		return nil
	}
	var b strings.Builder
	b.Grow(len(updated) + len(strings.Join(insert, "")) + 16)
	b.WriteString(updated[:idx])
	for _, item := range insert {
		b.WriteString(item)
	}
	b.WriteString(updated[idx:])
	files[contentTypesPath] = bytes.NewReader(StringToBytes(b.String()))
	return nil
}

type marshaller struct {
	data interface{}
	io.Reader
	io.WriterTo
}

// Read is fake and is to trigger io.WriterTo
func (m marshaller) Read(_ []byte) (int, error) {
	return 0, os.ErrInvalid
}

// WriteTo n is always 0 for we don't care that value
func (m marshaller) WriteTo(w io.Writer) (n int64, err error) {
	_, err = io.WriteString(w, xml.Header)
	if err != nil {
		return
	}
	err = xml.NewEncoder(w).Encode(m.data)
	return
}
