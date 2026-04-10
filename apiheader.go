package docx

import (
	"errors"
	"fmt"
	"strings"
)

var errNilHeaderFooter = errors.New("header/footer cannot be nil")
var errInvalidAlignment = errors.New("invalid alignment")
var errSectionOutOfRange = errors.New("section index out of range")
var errSettingsPartMissing = errors.New("settings.xml part missing")

// GetHeader returns header by kind.
func (f *Docx) GetHeader(kind HeaderKind) (*Header, error) {
	kind = normalizeHeaderKind(kind)
	if idx := f.mainSectionIndex(false); idx >= 0 {
		return f.GetSectionHeader(idx, kind)
	}
	f.ensureHeaderFooterMaps()
	return f.headers[kind], nil
}

// GetFooter returns footer by kind.
func (f *Docx) GetFooter(kind FooterKind) (*Footer, error) {
	kind = normalizeFooterKind(kind)
	if idx := f.mainSectionIndex(false); idx >= 0 {
		return f.GetSectionFooter(idx, kind)
	}
	f.ensureHeaderFooterMaps()
	return f.footers[kind], nil
}

// SetHeader sets header for kind.
func (f *Docx) SetHeader(kind HeaderKind, h *Header) error {
	idx := f.mainSectionIndex(true)
	if idx < 0 {
		return errSectionOutOfRange
	}
	return f.SetSectionHeader(idx, kind, h)
}

// SetFooter sets footer for kind.
func (f *Docx) SetFooter(kind FooterKind, ft *Footer) error {
	idx := f.mainSectionIndex(true)
	if idx < 0 {
		return errSectionOutOfRange
	}
	return f.SetSectionFooter(idx, kind, ft)
}

// NewHeader creates an empty header object.
func (f *Docx) NewHeader() *Header {
	return &Header{file: f}
}

// NewFooter creates an empty footer object.
func (f *Docx) NewFooter() *Footer {
	return &Footer{file: f}
}

// SectionCount returns total section count in document order.
func (f *Docx) SectionCount() int {
	return len(f.allSectionsInOrder())
}

// GetSectionHeader returns section-scoped header by kind.
func (f *Docx) GetSectionHeader(section int, kind HeaderKind) (*Header, error) {
	sect, err := f.sectionByIndexOrError(section)
	if err != nil {
		return nil, err
	}
	return f.getSectionHeaderObject(sect, kind), nil
}

// GetSectionFooter returns section-scoped footer by kind.
func (f *Docx) GetSectionFooter(section int, kind FooterKind) (*Footer, error) {
	sect, err := f.sectionByIndexOrError(section)
	if err != nil {
		return nil, err
	}
	return f.getSectionFooterObject(sect, kind), nil
}

// SetSectionHeader sets section-scoped header by kind.
func (f *Docx) SetSectionHeader(section int, kind HeaderKind, h *Header) error {
	if h == nil {
		return errNilHeaderFooter
	}
	sect, err := f.sectionByIndexOrError(section)
	if err != nil {
		return err
	}
	kind = normalizeHeaderKind(kind)
	f.setSectionHeaderObject(sect, kind, h)
	f.markSectionHeaderDirty(sect, kind)
	if section == f.mainSectionIndex(false) {
		f.ensureHeaderFooterMaps()
		f.headers[kind] = h
	}
	return nil
}

// SetSectionFooter sets section-scoped footer by kind.
func (f *Docx) SetSectionFooter(section int, kind FooterKind, ft *Footer) error {
	if ft == nil {
		return errNilHeaderFooter
	}
	sect, err := f.sectionByIndexOrError(section)
	if err != nil {
		return err
	}
	kind = normalizeFooterKind(kind)
	f.setSectionFooterObject(sect, kind, ft)
	f.markSectionFooterDirty(sect, kind)
	if section == f.mainSectionIndex(false) {
		f.ensureHeaderFooterMaps()
		f.footers[kind] = ft
	}
	return nil
}

// SetSectionHeaderText creates/replaces section header with a single text paragraph.
func (f *Docx) SetSectionHeaderText(section int, kind HeaderKind, text string) error {
	sect, err := f.sectionByIndexOrError(section)
	if err != nil {
		return err
	}
	kind = normalizeHeaderKind(kind)
	h, ok := f.sectionWritableHeader(sect, kind)
	if !ok || h == nil {
		h = f.NewHeader()
		f.setSectionHeaderObject(sect, kind, h)
	}
	h.Items = nil
	h.ordered = nil
	h.AddText(text)
	f.markSectionHeaderDirty(sect, kind)
	if section == f.mainSectionIndex(false) {
		f.ensureHeaderFooterMaps()
		f.headers[kind] = h
	}
	return nil
}

// SetSectionFooterText creates/replaces section footer with a single text paragraph.
func (f *Docx) SetSectionFooterText(section int, kind FooterKind, text string) error {
	sect, err := f.sectionByIndexOrError(section)
	if err != nil {
		return err
	}
	kind = normalizeFooterKind(kind)
	ft, ok := f.sectionWritableFooter(sect, kind)
	if !ok || ft == nil {
		ft = f.NewFooter()
		f.setSectionFooterObject(sect, kind, ft)
	}
	ft.Items = nil
	ft.ordered = nil
	ft.AddText(text)
	f.markSectionFooterDirty(sect, kind)
	if section == f.mainSectionIndex(false) {
		f.ensureHeaderFooterMaps()
		f.footers[kind] = ft
	}
	return nil
}

// SetSectionTitlePage toggles titlePg on a section.
func (f *Docx) SetSectionTitlePage(section int, enabled bool) error {
	sect, err := f.sectionByIndexOrError(section)
	if err != nil {
		return err
	}
	sect.setTitlePage(enabled)
	return nil
}

// SetEvenAndOddHeaders toggles evenAndOddHeaders in word/settings.xml.
// It only works when the source document already contains settings.xml.
func (f *Docx) SetEvenAndOddHeaders(enabled bool) error {
	if !f.settingsExists || f.settings == nil {
		return errSettingsPartMissing
	}
	f.settings.setEvenAndOddHeaders(enabled)
	f.settingsDirty = true
	return nil
}

// SetSectionHeaderAlignment sets top-level paragraph alignment in section header.
func (f *Docx) SetSectionHeaderAlignment(section int, kind HeaderKind, align string) error {
	normalized, err := normalizeParagraphAlignment(align)
	if err != nil {
		return err
	}
	sect, err := f.sectionByIndexOrError(section)
	if err != nil {
		return err
	}
	kind = normalizeHeaderKind(kind)
	h, ok := f.sectionWritableHeader(sect, kind)
	if !ok || h == nil {
		h = f.NewHeader()
		f.setSectionHeaderObject(sect, kind, h)
	}
	if !applyAlignmentToTopLevelParagraphs(h.Items, normalized) {
		h.AddParagraph().Justification(normalized)
	}
	f.markSectionHeaderDirty(sect, kind)
	if section == f.mainSectionIndex(false) {
		f.ensureHeaderFooterMaps()
		f.headers[kind] = h
	}
	return nil
}

// SetSectionFooterAlignment sets top-level paragraph alignment in section footer.
func (f *Docx) SetSectionFooterAlignment(section int, kind FooterKind, align string) error {
	normalized, err := normalizeParagraphAlignment(align)
	if err != nil {
		return err
	}
	sect, err := f.sectionByIndexOrError(section)
	if err != nil {
		return err
	}
	kind = normalizeFooterKind(kind)
	ft, ok := f.sectionWritableFooter(sect, kind)
	if !ok || ft == nil {
		ft = f.NewFooter()
		f.setSectionFooterObject(sect, kind, ft)
	}
	if !applyAlignmentToTopLevelParagraphs(ft.Items, normalized) {
		ft.AddParagraph().Justification(normalized)
	}
	f.markSectionFooterDirty(sect, kind)
	if section == f.mainSectionIndex(false) {
		f.ensureHeaderFooterMaps()
		f.footers[kind] = ft
	}
	return nil
}

// AddSectionPageNumber appends PAGE field to section footer.
func (f *Docx) AddSectionPageNumber(section int, style PageNumberStyle, kind ...FooterKind) error {
	target := FooterDefault
	if len(kind) > 0 {
		target = normalizeFooterKind(kind[0])
	}
	_, err := f.addPageNumberToSection(section, style, target)
	return err
}

// AddSectionPageNumberAligned appends PAGE field and aligns the inserted paragraph.
func (f *Docx) AddSectionPageNumberAligned(section int, style PageNumberStyle, align string, kind ...FooterKind) error {
	normalized, err := normalizeParagraphAlignment(align)
	if err != nil {
		return err
	}
	target := FooterDefault
	if len(kind) > 0 {
		target = normalizeFooterKind(kind[0])
	}
	p, err := f.addPageNumberToSection(section, style, target)
	if err != nil {
		return err
	}
	p.Justification(normalized)
	return nil
}

// SetHeaderText creates/replaces a header with a single text paragraph.
func (f *Docx) SetHeaderText(kind HeaderKind, text string) error {
	idx := f.mainSectionIndex(true)
	if idx < 0 {
		return errSectionOutOfRange
	}
	return f.SetSectionHeaderText(idx, kind, text)
}

// SetFooterText creates/replaces a footer with a single text paragraph.
func (f *Docx) SetFooterText(kind FooterKind, text string) error {
	idx := f.mainSectionIndex(true)
	if idx < 0 {
		return errSectionOutOfRange
	}
	return f.SetSectionFooterText(idx, kind, text)
}

// AddPageNumber appends PAGE field to target footer.
// If kind is omitted, FooterDefault is used.
func (f *Docx) AddPageNumber(style PageNumberStyle, kind ...FooterKind) error {
	target := FooterDefault
	if len(kind) > 0 {
		target = normalizeFooterKind(kind[0])
	}
	_, err := f.addPageNumberToFooter(style, target)
	return err
}

// AddPageNumberAligned appends PAGE field and aligns the paragraph where
// the field is inserted.
func (f *Docx) AddPageNumberAligned(style PageNumberStyle, align string, kind ...FooterKind) error {
	normalized, err := normalizeParagraphAlignment(align)
	if err != nil {
		return err
	}
	target := FooterDefault
	if len(kind) > 0 {
		target = normalizeFooterKind(kind[0])
	}
	p, err := f.addPageNumberToFooter(style, target)
	if err != nil {
		return err
	}
	p.Justification(normalized)
	return nil
}

// AddParagraph appends a paragraph to header.
func (h *Header) AddParagraph() *Paragraph {
	p := &Paragraph{file: h.file}
	h.Items = append(h.Items, p)
	h.ordered = append(h.ordered, p)
	return p
}

// AddText appends a paragraph with one text run to header.
func (h *Header) AddText(text string) *Run {
	return h.AddParagraph().AddText(text)
}

// AddParagraph appends a paragraph to footer.
func (f *Footer) AddParagraph() *Paragraph {
	p := &Paragraph{file: f.file}
	f.Items = append(f.Items, p)
	f.ordered = append(f.ordered, p)
	return p
}

// AddText appends a paragraph with one text run to footer.
func (f *Footer) AddText(text string) *Run {
	return f.AddParagraph().AddText(text)
}

// LastParagraph returns the last paragraph from footer if exists.
func (f *Footer) LastParagraph() *Paragraph {
	for i := len(f.Items) - 1; i >= 0; i-- {
		if p, ok := f.Items[i].(*Paragraph); ok {
			return p
		}
	}
	return nil
}

// SetHeaderAlignment sets paragraph alignment for all top-level paragraphs
// in a header. If no paragraph exists, one is created.
func (f *Docx) SetHeaderAlignment(kind HeaderKind, align string) error {
	idx := f.mainSectionIndex(true)
	if idx < 0 {
		return errSectionOutOfRange
	}
	return f.SetSectionHeaderAlignment(idx, kind, align)
}

// SetFooterAlignment sets paragraph alignment for all top-level paragraphs
// in a footer. If no paragraph exists, one is created.
func (f *Docx) SetFooterAlignment(kind FooterKind, align string) error {
	idx := f.mainSectionIndex(true)
	if idx < 0 {
		return errSectionOutOfRange
	}
	return f.SetSectionFooterAlignment(idx, kind, align)
}

func (f *Docx) addPageNumberToFooter(style PageNumberStyle, target FooterKind) (*Paragraph, error) {
	idx := f.mainSectionIndex(true)
	if idx < 0 {
		return nil, errSectionOutOfRange
	}
	return f.addPageNumberToSection(idx, style, target)
}

func (f *Docx) addPageNumberToSection(section int, style PageNumberStyle, target FooterKind) (*Paragraph, error) {
	sect, err := f.sectionByIndexOrError(section)
	if err != nil {
		return nil, err
	}
	target = normalizeFooterKind(target)
	ft, ok := f.sectionWritableFooter(sect, target)
	if !ok || ft == nil {
		ft = f.NewFooter()
		f.setSectionFooterObject(sect, target, ft)
	}
	p := ft.LastParagraph()
	if p == nil {
		p = ft.AddParagraph()
	}
	addPageNumberField(p, style)
	sect.setPageNumberFormat(pageNumberFormat(style))
	f.markSectionFooterDirty(sect, target)
	if section == f.mainSectionIndex(false) {
		f.ensureHeaderFooterMaps()
		f.footers[target] = ft
	}
	return p, nil
}

func normalizeParagraphAlignment(align string) (string, error) {
	a := strings.ToLower(strings.TrimSpace(align))
	switch a {
	case "start", "center", "end", "both", "distribute":
		return a, nil
	default:
		return "", fmt.Errorf("%w: %q", errInvalidAlignment, align)
	}
}

func applyAlignmentToTopLevelParagraphs(items []interface{}, align string) bool {
	applied := false
	for _, item := range items {
		p, ok := item.(*Paragraph)
		if !ok {
			continue
		}
		p.Justification(align)
		applied = true
	}
	return applied
}

func pageNumberFormat(style PageNumberStyle) string {
	switch style {
	case PageNumberRoman:
		return "lowerRoman"
	case PageNumberRomanUpper:
		return "upperRoman"
	case PageNumberLetter:
		return "lowerLetter"
	case PageNumberLetterUpper:
		return "upperLetter"
	default:
		return "decimal"
	}
}

func addPageNumberField(p *Paragraph, style PageNumberStyle) {
	if len(p.ordered) == 0 && len(p.Children) > 0 {
		p.ordered = append([]interface{}{}, p.Children...)
	}
	rBegin := &Run{
		Children: []interface{}{
			&runFldChar{Type: "begin"},
		},
	}
	instr := " PAGE "
	switch style {
	case PageNumberRoman:
		instr = " PAGE \\* roman "
	case PageNumberRomanUpper:
		instr = " PAGE \\* ROMAN "
	case PageNumberLetter:
		instr = " PAGE \\* letter "
	case PageNumberLetterUpper:
		instr = " PAGE \\* LETTER "
	}
	rInstr := &Run{
		InstrText: instr,
		ordered:   []interface{}{&runInstrText{Text: instr}},
	}
	rSep := &Run{
		Children: []interface{}{
			&runFldChar{Type: "separate"},
		},
	}
	rEnd := &Run{
		Children: []interface{}{
			&runFldChar{Type: "end"},
		},
	}
	p.Children = append(p.Children, rBegin, rInstr, rSep, rEnd)
	p.ordered = append(p.ordered, rBegin, rInstr, rSep, rEnd)
}

func (f *Docx) sectionByIndexOrError(section int) (*SectPr, error) {
	sect := f.sectionByIndex(section)
	if sect == nil {
		return nil, fmt.Errorf("%w: %d", errSectionOutOfRange, section)
	}
	return sect, nil
}

func normalizeHeaderKind(kind HeaderKind) HeaderKind {
	k := strings.ToLower(strings.TrimSpace(string(kind)))
	switch HeaderKind(k) {
	case HeaderFirst:
		return HeaderFirst
	case HeaderEven:
		return HeaderEven
	default:
		return HeaderDefault
	}
}

func normalizeFooterKind(kind FooterKind) FooterKind {
	k := strings.ToLower(strings.TrimSpace(string(kind)))
	switch FooterKind(k) {
	case FooterFirst:
		return FooterFirst
	case FooterEven:
		return FooterEven
	default:
		return FooterDefault
	}
}

func (f *Docx) ensureHeaderFooterMaps() {
	if f.headers == nil {
		f.headers = make(map[HeaderKind]*Header, 3)
	}
	if f.footers == nil {
		f.footers = make(map[FooterKind]*Footer, 3)
	}
}
