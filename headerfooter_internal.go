package docx

import (
	"bytes"
	"encoding/xml"
	"strings"
)

type sectionHeaderFooterRefs struct {
	headers map[HeaderKind]*Header
	footers map[FooterKind]*Footer
}

func (f *Docx) ensureMainSectPr(create bool) *SectPr {
	var main *SectPr
	for i := len(f.Document.Body.Items) - 1; i >= 0; i-- {
		if s, ok := f.Document.Body.Items[i].(*SectPr); ok {
			main = s
			break
		}
	}
	if main != nil || !create {
		return main
	}
	main = &SectPr{}
	f.Document.Body.Items = append(f.Document.Body.Items, main)
	return main
}

func (f *Docx) appendBodyItemBeforeTrailingSectPr(item interface{}) {
	items := f.Document.Body.Items
	n := len(items)
	if n > 0 {
		if _, ok := items[n-1].(*SectPr); ok {
			f.Document.Body.Items = append(items[:n-1], item, items[n-1])
			return
		}
	}
	f.Document.Body.Items = append(items, item)
}

func headerKindsInOrder() []HeaderKind {
	return []HeaderKind{HeaderDefault, HeaderFirst, HeaderEven}
}

func footerKindsInOrder() []FooterKind {
	return []FooterKind{FooterDefault, FooterFirst, FooterEven}
}

func headerKindFromRefType(t string) (HeaderKind, bool) {
	switch HeaderKind(strings.ToLower(strings.TrimSpace(t))) {
	case HeaderDefault:
		return HeaderDefault, true
	case HeaderFirst:
		return HeaderFirst, true
	case HeaderEven:
		return HeaderEven, true
	default:
		return "", false
	}
}

func footerKindFromRefType(t string) (FooterKind, bool) {
	switch FooterKind(strings.ToLower(strings.TrimSpace(t))) {
	case FooterDefault:
		return FooterDefault, true
	case FooterFirst:
		return FooterFirst, true
	case FooterEven:
		return FooterEven, true
	default:
		return "", false
	}
}

func (s *SectPr) setHeaderFooterRefs(hrefs []*HeaderReference, frefs []*FooterReference) {
	s.HeaderRefs = hrefs
	s.FooterRefs = frefs
	if len(s.ordered) == 0 {
		return
	}
	rest := make([]interface{}, 0, len(s.ordered))
	for _, item := range s.ordered {
		switch item.(type) {
		case *HeaderReference, *FooterReference:
			continue
		default:
			rest = append(rest, item)
		}
	}
	newOrdered := make([]interface{}, 0, len(hrefs)+len(frefs)+len(rest))
	for _, r := range hrefs {
		newOrdered = append(newOrdered, r)
	}
	for _, r := range frefs {
		newOrdered = append(newOrdered, r)
	}
	newOrdered = append(newOrdered, rest...)
	s.ordered = newOrdered
}

func (s *SectPr) setPageNumberFormat(fmtValue string) {
	fmtValue = strings.TrimSpace(fmtValue)
	if fmtValue == "" {
		return
	}
	if s.PgNumType != nil {
		s.PgNumType.Fmt = fmtValue
		return
	}
	s.PgNumType = &PgNumType{Fmt: fmtValue}
	if len(s.ordered) == 0 {
		return
	}
	insertAt := len(s.ordered)
	for i, item := range s.ordered {
		switch item.(type) {
		case *PgMar, *Cols, *DocGrid:
			insertAt = i
			goto insert
		}
	}
insert:
	s.ordered = append(s.ordered, nil)
	copy(s.ordered[insertAt+1:], s.ordered[insertAt:])
	s.ordered[insertAt] = s.PgNumType
}

func (s *SectPr) setTitlePage(enabled bool) {
	if enabled {
		if s.TitlePg != nil {
			s.TitlePg.Val = ""
			return
		}
		s.TitlePg = &OnOff{}
		if len(s.ordered) == 0 {
			return
		}
		insertAt := len(s.ordered)
		for i, item := range s.ordered {
			switch item.(type) {
			case *PgSz, *PgNumType, *PgMar, *Cols, *DocGrid:
				insertAt = i
				goto insert
			}
		}
	insert:
		s.ordered = append(s.ordered, nil)
		copy(s.ordered[insertAt+1:], s.ordered[insertAt:])
		s.ordered[insertAt] = s.TitlePg
		return
	}
	if s.TitlePg == nil {
		return
	}
	toRemove := s.TitlePg
	s.TitlePg = nil
	if len(s.ordered) == 0 {
		return
	}
	next := make([]interface{}, 0, len(s.ordered))
	for _, item := range s.ordered {
		if item == toRemove {
			continue
		}
		next = append(next, item)
	}
	s.ordered = next
}

func (f *Docx) ensureSectionRefMaps() {
	if f.sectionRefs == nil {
		f.sectionRefs = make(map[*SectPr]*sectionHeaderFooterRefs, 4)
	}
	if f.sectionHeaderDirty == nil {
		f.sectionHeaderDirty = make(map[*SectPr]map[HeaderKind]bool, 4)
	}
	if f.sectionFooterDirty == nil {
		f.sectionFooterDirty = make(map[*SectPr]map[FooterKind]bool, 4)
	}
}

func (f *Docx) allSectionsInOrder() []*SectPr {
	sections := make([]*SectPr, 0, 4)
	for _, item := range f.Document.Body.Items {
		switch v := item.(type) {
		case *Paragraph:
			if v.Properties != nil && v.Properties.SectPr != nil {
				sections = append(sections, v.Properties.SectPr)
			}
		case *SectPr:
			sections = append(sections, v)
		}
	}
	return sections
}

func (f *Docx) sectionIndexOf(target *SectPr) int {
	if target == nil {
		return -1
	}
	sections := f.allSectionsInOrder()
	for i, s := range sections {
		if s == target {
			return i
		}
	}
	return -1
}

func (f *Docx) mainSectionIndex(create bool) int {
	main := f.ensureMainSectPr(create)
	if main == nil {
		return -1
	}
	return f.sectionIndexOf(main)
}

func (f *Docx) sectionByIndex(section int) *SectPr {
	if section < 0 {
		return nil
	}
	sections := f.allSectionsInOrder()
	if section >= len(sections) {
		return nil
	}
	return sections[section]
}

func (f *Docx) sectionRefsFor(sect *SectPr, create bool) *sectionHeaderFooterRefs {
	if sect == nil {
		return nil
	}
	f.ensureSectionRefMaps()
	if refs, ok := f.sectionRefs[sect]; ok {
		return refs
	}
	if !create {
		return nil
	}
	refs := &sectionHeaderFooterRefs{
		headers: make(map[HeaderKind]*Header, 3),
		footers: make(map[FooterKind]*Footer, 3),
	}
	f.sectionRefs[sect] = refs
	return refs
}

func (f *Docx) getSectionHeaderObject(sect *SectPr, kind HeaderKind) *Header {
	refs := f.sectionRefsFor(sect, false)
	if refs == nil {
		return nil
	}
	return refs.headers[normalizeHeaderKind(kind)]
}

func (f *Docx) getSectionFooterObject(sect *SectPr, kind FooterKind) *Footer {
	refs := f.sectionRefsFor(sect, false)
	if refs == nil {
		return nil
	}
	return refs.footers[normalizeFooterKind(kind)]
}

func (f *Docx) setSectionHeaderObject(sect *SectPr, kind HeaderKind, h *Header) {
	refs := f.sectionRefsFor(sect, true)
	if refs.headers == nil {
		refs.headers = make(map[HeaderKind]*Header, 3)
	}
	kind = normalizeHeaderKind(kind)
	if h == nil {
		delete(refs.headers, kind)
		return
	}
	h.file = f
	refs.headers[kind] = h
}

func (f *Docx) setSectionFooterObject(sect *SectPr, kind FooterKind, ft *Footer) {
	refs := f.sectionRefsFor(sect, true)
	if refs.footers == nil {
		refs.footers = make(map[FooterKind]*Footer, 3)
	}
	kind = normalizeFooterKind(kind)
	if ft == nil {
		delete(refs.footers, kind)
		return
	}
	ft.file = f
	refs.footers[kind] = ft
}

func (f *Docx) markSectionHeaderDirty(sect *SectPr, kind HeaderKind) {
	if sect == nil {
		return
	}
	f.ensureSectionRefMaps()
	if f.sectionHeaderDirty[sect] == nil {
		f.sectionHeaderDirty[sect] = make(map[HeaderKind]bool, 3)
	}
	f.sectionHeaderDirty[sect][normalizeHeaderKind(kind)] = true
}

func (f *Docx) markSectionFooterDirty(sect *SectPr, kind FooterKind) {
	if sect == nil {
		return
	}
	f.ensureSectionRefMaps()
	if f.sectionFooterDirty[sect] == nil {
		f.sectionFooterDirty[sect] = make(map[FooterKind]bool, 3)
	}
	f.sectionFooterDirty[sect][normalizeFooterKind(kind)] = true
}

func (f *Docx) isSectionHeaderDirty(sect *SectPr, kind HeaderKind) bool {
	if sect == nil || f.sectionHeaderDirty == nil {
		return false
	}
	m := f.sectionHeaderDirty[sect]
	if m == nil {
		return false
	}
	return m[normalizeHeaderKind(kind)]
}

func (f *Docx) isSectionFooterDirty(sect *SectPr, kind FooterKind) bool {
	if sect == nil || f.sectionFooterDirty == nil {
		return false
	}
	m := f.sectionFooterDirty[sect]
	if m == nil {
		return false
	}
	return m[normalizeFooterKind(kind)]
}

func (f *Docx) syncLegacyMainSectionMaps() {
	f.ensureHeaderFooterMaps()
	for k := range f.headers {
		delete(f.headers, k)
	}
	for k := range f.footers {
		delete(f.footers, k)
	}
	main := f.ensureMainSectPr(false)
	if main == nil {
		return
	}
	refs := f.sectionRefsFor(main, false)
	if refs == nil {
		return
	}
	for _, kind := range headerKindsInOrder() {
		if h := refs.headers[kind]; h != nil {
			f.headers[kind] = h
		}
	}
	for _, kind := range footerKindsInOrder() {
		if ft := refs.footers[kind]; ft != nil {
			f.footers[kind] = ft
		}
	}
}

func (f *Docx) syncMainSectionFromLegacyMaps() {
	if len(f.headers) == 0 && len(f.footers) == 0 {
		return
	}
	main := f.ensureMainSectPr(true)
	for _, kind := range headerKindsInOrder() {
		if h := f.headers[kind]; h != nil {
			f.setSectionHeaderObject(main, kind, h)
		}
	}
	for _, kind := range footerKindsInOrder() {
		if ft := f.footers[kind]; ft != nil {
			f.setSectionFooterObject(main, kind, ft)
		}
	}
}

func (f *Docx) sectionWritableHeader(sect *SectPr, kind HeaderKind) (*Header, bool) {
	h := f.getSectionHeaderObject(sect, kind)
	if h == nil {
		return nil, false
	}
	if f.countHeaderUsage(h) > 1 {
		cloned, err := cloneHeader(h, f)
		if err == nil {
			f.setSectionHeaderObject(sect, kind, cloned)
			h = cloned
		}
	}
	return h, true
}

func (f *Docx) sectionWritableFooter(sect *SectPr, kind FooterKind) (*Footer, bool) {
	ft := f.getSectionFooterObject(sect, kind)
	if ft == nil {
		return nil, false
	}
	if f.countFooterUsage(ft) > 1 {
		cloned, err := cloneFooter(ft, f)
		if err == nil {
			f.setSectionFooterObject(sect, kind, cloned)
			ft = cloned
		}
	}
	return ft, true
}

func (f *Docx) countHeaderUsage(target *Header) int {
	if target == nil {
		return 0
	}
	count := 0
	for _, refs := range f.sectionRefs {
		for _, h := range refs.headers {
			if h == target {
				count++
			}
		}
	}
	return count
}

func (f *Docx) countFooterUsage(target *Footer) int {
	if target == nil {
		return 0
	}
	count := 0
	for _, refs := range f.sectionRefs {
		for _, ft := range refs.footers {
			if ft == target {
				count++
			}
		}
	}
	return count
}

func cloneHeader(src *Header, file *Docx) (*Header, error) {
	var buf bytes.Buffer
	if err := xml.NewEncoder(&buf).Encode(src); err != nil {
		return nil, err
	}
	var out Header
	out.file = file
	if err := xml.NewDecoder(bytes.NewReader(buf.Bytes())).Decode(&out); err != nil {
		return nil, err
	}
	return &out, nil
}

func cloneFooter(src *Footer, file *Docx) (*Footer, error) {
	var buf bytes.Buffer
	if err := xml.NewEncoder(&buf).Encode(src); err != nil {
		return nil, err
	}
	var out Footer
	out.file = file
	if err := xml.NewDecoder(bytes.NewReader(buf.Bytes())).Decode(&out); err != nil {
		return nil, err
	}
	return &out, nil
}
