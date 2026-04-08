package docx

import (
	"errors"
	"strings"
)

var errNilHeaderFooter = errors.New("header/footer cannot be nil")

// GetHeader returns header by kind.
func (f *Docx) GetHeader(kind HeaderKind) (*Header, error) {
	f.ensureHeaderFooterMaps()
	kind = normalizeHeaderKind(kind)
	return f.headers[kind], nil
}

// GetFooter returns footer by kind.
func (f *Docx) GetFooter(kind FooterKind) (*Footer, error) {
	f.ensureHeaderFooterMaps()
	kind = normalizeFooterKind(kind)
	return f.footers[kind], nil
}

// SetHeader sets header for kind.
func (f *Docx) SetHeader(kind HeaderKind, h *Header) error {
	if h == nil {
		return errNilHeaderFooter
	}
	f.ensureHeaderFooterMaps()
	kind = normalizeHeaderKind(kind)
	h.file = f
	f.headers[kind] = h
	f.ensureMainSectPr(true)
	return nil
}

// SetFooter sets footer for kind.
func (f *Docx) SetFooter(kind FooterKind, ft *Footer) error {
	if ft == nil {
		return errNilHeaderFooter
	}
	f.ensureHeaderFooterMaps()
	kind = normalizeFooterKind(kind)
	ft.file = f
	f.footers[kind] = ft
	f.ensureMainSectPr(true)
	return nil
}

// NewHeader creates an empty header object.
func (f *Docx) NewHeader() *Header {
	return &Header{file: f}
}

// NewFooter creates an empty footer object.
func (f *Docx) NewFooter() *Footer {
	return &Footer{file: f}
}

// SetHeaderText creates/replaces a header with a single text paragraph.
func (f *Docx) SetHeaderText(kind HeaderKind, text string) error {
	h := f.NewHeader()
	h.AddText(text)
	return f.SetHeader(kind, h)
}

// SetFooterText creates/replaces a footer with a single text paragraph.
func (f *Docx) SetFooterText(kind FooterKind, text string) error {
	ft := f.NewFooter()
	ft.AddText(text)
	return f.SetFooter(kind, ft)
}

// AddPageNumber appends PAGE field to target footer.
// If kind is omitted, FooterDefault is used.
func (f *Docx) AddPageNumber(style PageNumberStyle, kind ...FooterKind) error {
	target := FooterDefault
	if len(kind) > 0 {
		target = normalizeFooterKind(kind[0])
	}
	ft, _ := f.GetFooter(target)
	if ft == nil {
		ft = f.NewFooter()
		if err := f.SetFooter(target, ft); err != nil {
			return err
		}
	}
	p := ft.LastParagraph()
	if p == nil {
		p = ft.AddParagraph()
	}
	addPageNumberField(p, style)
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
