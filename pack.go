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
	"crypto/sha256"
	"encoding/hex"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"sort"
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
	if f.settingsExists && f.settingsDirty && f.settings != nil {
		files["word/settings.xml"] = marshaller{data: f.settings}
	}

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
	f.syncMainSectionFromLegacyMaps()
	sections := f.allSectionsInOrder()
	if len(sections) == 0 && (len(f.headers) > 0 || len(f.footers) > 0 || len(f.sectionRefs) > 0) {
		_ = f.ensureMainSectPr(true)
		sections = f.allSectionsInOrder()
	}
	if len(sections) == 0 {
		return nil, nil, nil
	}
	headerPathSet := make(map[string]struct{}, 8)
	footerPathSet := make(map[string]struct{}, 8)
	headerDedup := make(map[string]string, 8)
	footerDedup := make(map[string]string, 8)

	singleSection := len(sections) == 1
	for sectionIndex, sect := range sections {
		existingHeaderKnown := make(map[HeaderKind]*HeaderReference, 3)
		existingFooterKnown := make(map[FooterKind]*FooterReference, 3)
		existingHeaderUnknown := make([]*HeaderReference, 0, len(sect.HeaderRefs))
		existingFooterUnknown := make([]*FooterReference, 0, len(sect.FooterRefs))

		for _, ref := range sect.HeaderRefs {
			if ref == nil {
				continue
			}
			if kind, ok := headerKindFromRefType(ref.Type); ok {
				existingHeaderKnown[kind] = ref
			} else {
				existingHeaderUnknown = append(existingHeaderUnknown, ref)
			}
		}
		for _, ref := range sect.FooterRefs {
			if ref == nil {
				continue
			}
			if kind, ok := footerKindFromRefType(ref.Type); ok {
				existingFooterKnown[kind] = ref
			} else {
				existingFooterUnknown = append(existingFooterUnknown, ref)
			}
		}

		hrefs := make([]*HeaderReference, 0, len(existingHeaderKnown)+len(existingHeaderUnknown))
		frefs := make([]*FooterReference, 0, len(existingFooterKnown)+len(existingFooterUnknown))

		for _, kind := range headerKindsInOrder() {
			h := f.getSectionHeaderObject(sect, kind)
			existing := existingHeaderKnown[kind]
			if h == nil {
				if existing != nil {
					hrefs = append(hrefs, existing)
				}
				continue
			}
			h.file = f
			sig, err := xmlPartSignature(h)
			if err != nil {
				return nil, nil, err
			}
			path := ""
			currentRID := ""
			if existing != nil {
				currentRID = existing.RID
				if !f.isSectionHeaderDirty(sect, kind) {
					if rel := f.findRelationshipByID(existing.RID); rel != nil {
						path = "word/" + normalizeRelTarget(rel.Target)
					}
				}
			}
			if target, ok := headerDedup[sig]; ok {
				rid := f.ensureInternalPartRelation(REL_HEADER, target, currentRID)
				hrefs = append(hrefs, &HeaderReference{Type: string(kind), RID: rid})
				continue
			}
			if path == "" {
				path = defaultHeaderPartPath(sectionIndex, kind, singleSection)
				currentRID = ""
			}
			target := strings.TrimPrefix(path, "word/")
			files[path] = marshaller{data: h}
			headerPathSet[path] = struct{}{}
			rid := f.ensureInternalPartRelation(REL_HEADER, target, currentRID)
			hrefs = append(hrefs, &HeaderReference{Type: string(kind), RID: rid})
			headerDedup[sig] = target
		}

		for _, kind := range footerKindsInOrder() {
			ft := f.getSectionFooterObject(sect, kind)
			existing := existingFooterKnown[kind]
			if ft == nil {
				if existing != nil {
					frefs = append(frefs, existing)
				}
				continue
			}
			ft.file = f
			sig, err := xmlPartSignature(ft)
			if err != nil {
				return nil, nil, err
			}
			path := ""
			currentRID := ""
			if existing != nil {
				currentRID = existing.RID
				if !f.isSectionFooterDirty(sect, kind) {
					if rel := f.findRelationshipByID(existing.RID); rel != nil {
						path = "word/" + normalizeRelTarget(rel.Target)
					}
				}
			}
			if target, ok := footerDedup[sig]; ok {
				rid := f.ensureInternalPartRelation(REL_FOOTER, target, currentRID)
				frefs = append(frefs, &FooterReference{Type: string(kind), RID: rid})
				continue
			}
			if path == "" {
				path = defaultFooterPartPath(sectionIndex, kind, singleSection)
				currentRID = ""
			}
			target := strings.TrimPrefix(path, "word/")
			files[path] = marshaller{data: ft}
			footerPathSet[path] = struct{}{}
			rid := f.ensureInternalPartRelation(REL_FOOTER, target, currentRID)
			frefs = append(frefs, &FooterReference{Type: string(kind), RID: rid})
			footerDedup[sig] = target
		}

		hrefs = append(hrefs, existingHeaderUnknown...)
		frefs = append(frefs, existingFooterUnknown...)
		sect.setHeaderFooterRefs(hrefs, frefs)
	}

	headerPaths := make([]string, 0, len(headerPathSet))
	for p := range headerPathSet {
		headerPaths = append(headerPaths, p)
	}
	footerPaths := make([]string, 0, len(footerPathSet))
	for p := range footerPathSet {
		footerPaths = append(footerPaths, p)
	}
	sort.Strings(headerPaths)
	sort.Strings(footerPaths)
	return headerPaths, footerPaths, nil
}

func xmlPartSignature(v interface{}) (string, error) {
	var raw bytes.Buffer
	enc := xml.NewEncoder(&raw)
	if err := enc.Encode(v); err != nil {
		return "", err
	}
	if err := enc.Flush(); err != nil {
		return "", err
	}
	canonical, err := canonicalizeXML(raw.Bytes())
	if err != nil {
		return "", err
	}
	sum := sha256.Sum256(canonical)
	return hex.EncodeToString(sum[:]), nil
}

func canonicalizeXML(data []byte) ([]byte, error) {
	var out bytes.Buffer
	enc := xml.NewEncoder(&out)
	dec := xml.NewDecoder(bytes.NewReader(data))
	for {
		tok, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, err
		}
		switch v := tok.(type) {
		case xml.StartElement:
			v.Attr = normalizeAndSortAttrs(v.Attr)
			if err := enc.EncodeToken(v); err != nil {
				return nil, err
			}
		case xml.EndElement:
			if err := enc.EncodeToken(v); err != nil {
				return nil, err
			}
		case xml.CharData:
			if err := enc.EncodeToken(v); err != nil {
				return nil, err
			}
		case xml.Comment:
			if err := enc.EncodeToken(v); err != nil {
				return nil, err
			}
		case xml.ProcInst:
			if err := enc.EncodeToken(v); err != nil {
				return nil, err
			}
		case xml.Directive:
			if err := enc.EncodeToken(v); err != nil {
				return nil, err
			}
		}
	}
	if err := enc.Flush(); err != nil {
		return nil, err
	}
	return out.Bytes(), nil
}

func normalizeAndSortAttrs(attrs []xml.Attr) []xml.Attr {
	out := make([]xml.Attr, 0, len(attrs))
	for _, attr := range attrs {
		out = append(out, normalizeRootAttr(attr))
	}
	sort.Slice(out, func(i, j int) bool {
		ni := attrSortKey(out[i])
		nj := attrSortKey(out[j])
		if ni == nj {
			return out[i].Value < out[j].Value
		}
		return ni < nj
	})
	return out
}

func attrSortKey(attr xml.Attr) string {
	if attr.Name.Space == "" {
		return attr.Name.Local
	}
	return attr.Name.Space + ":" + attr.Name.Local
}

func defaultHeaderPartPath(sectionIndex int, kind HeaderKind, singleSection bool) string {
	if singleSection {
		return fmt.Sprintf("word/header_%s.xml", kind)
	}
	return fmt.Sprintf("word/header_s%d_%s.xml", sectionIndex+1, kind)
}

func defaultFooterPartPath(sectionIndex int, kind FooterKind, singleSection bool) string {
	if singleSection {
		return fmt.Sprintf("word/footer_%s.xml", kind)
	}
	return fmt.Sprintf("word/footer_s%d_%s.xml", sectionIndex+1, kind)
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
