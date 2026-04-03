package docx

import (
	"bytes"
	"encoding/xml"
	"os"
	"testing"
)

func TestReplaceTextSingleRun(t *testing.T) {
	d := New().WithDefaultTheme()
	p := d.AddParagraph()
	p.AddText("hello world")

	if err := d.ReplaceText("world", "docx"); err != nil {
		t.Fatal(err)
	}
	if got := p.String(); got != "hello docx" {
		t.Fatalf("unexpected paragraph text: %q", got)
	}
}

func TestReplaceTextCrossRuns(t *testing.T) {
	d := New().WithDefaultTheme()
	p := d.AddParagraph()
	p.AddText("hel")
	p.AddText("lo")
	p.AddText(" world")

	if err := d.ReplaceText("hello", "hi"); err != nil {
		t.Fatal(err)
	}
	if got := p.String(); got != "hi world" {
		t.Fatalf("unexpected paragraph text: %q", got)
	}
}

func TestReplaceTextWithMaxReplacements(t *testing.T) {
	d := New().WithDefaultTheme()
	p := d.AddParagraph()
	p.AddText("a a a")

	if err := d.ReplaceText("a", "b", WithMaxReplacements(2)); err != nil {
		t.Fatal(err)
	}
	if got := p.String(); got != "b b a" {
		t.Fatalf("unexpected paragraph text: %q", got)
	}
}

func TestReplaceTextCaseSensitivity(t *testing.T) {
	d := New().WithDefaultTheme()
	p := d.AddParagraph()
	p.AddText("Hello hello")

	if err := d.ReplaceText("hello", "x"); err != nil {
		t.Fatal(err)
	}
	if got := p.String(); got != "Hello x" {
		t.Fatalf("unexpected case-sensitive result: %q", got)
	}

	if err := d.ReplaceText("hello", "y", WithCaseSensitive(false)); err != nil {
		t.Fatal(err)
	}
	if got := p.String(); got != "y x" {
		t.Fatalf("unexpected case-insensitive result: %q", got)
	}
}

func TestReplacePlaceholder(t *testing.T) {
	d := New().WithDefaultTheme()
	p := d.AddParagraph()
	p.AddText("hello {{na")
	p.AddText("me}}")

	if err := d.ReplacePlaceholder(map[string]string{"name": "A"}); err != nil {
		t.Fatal(err)
	}
	if got := p.String(); got != "hello A" {
		t.Fatalf("unexpected placeholder result: %q", got)
	}
}

func TestReplaceTextInTableCellParagraph(t *testing.T) {
	d := New().WithDefaultTheme()
	tbl := d.AddTable(1, 1, 0, nil)
	p := tbl.TableRows[0].TableCells[0].AddParagraph()
	p.AddText("ab")
	p.AddText("cd")

	if err := d.ReplaceText("abcd", "ok"); err != nil {
		t.Fatal(err)
	}
	if got := p.String(); got != "ok" {
		t.Fatalf("unexpected table paragraph text: %q", got)
	}
}

func TestReplaceTextInHyperlinkRun(t *testing.T) {
	const in = `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:hyperlink r:id="rId1"><w:r><w:t>hello world</w:t></w:r></w:hyperlink></w:p>`
	var p Paragraph
	if err := xml.Unmarshal(StringToBytes(in), &p); err != nil {
		t.Fatal(err)
	}
	d := New().WithDefaultTheme()
	d.Document.Body.Items = append(d.Document.Body.Items, &p)

	if err := d.ReplaceText("world", "docx"); err != nil {
		t.Fatal(err)
	}

	h, ok := d.Document.Body.Items[0].(*Paragraph).Children[0].(*Hyperlink)
	if !ok {
		t.Fatalf("expected first child to be hyperlink, got %T", d.Document.Body.Items[0].(*Paragraph).Children[0])
	}
	if len(h.Run.Children) != 1 {
		t.Fatalf("expected one text child in hyperlink run, got %d", len(h.Run.Children))
	}
	txt, ok := h.Run.Children[0].(*Text)
	if !ok || txt.Text != "hello docx" {
		t.Fatalf("unexpected hyperlink text child: %#v", h.Run.Children[0])
	}
	if h.ID != "rId1" {
		t.Fatalf("hyperlink relation id changed: %q", h.ID)
	}
}

func TestReplaceTextDoesNotProcessInstrText(t *testing.T) {
	d := New().WithDefaultTheme()
	p := d.AddParagraph()
	p.Children = append(p.Children, &Run{
		RunProperties: &RunProperties{},
		InstrText:     "FORMTEXT hello",
	})

	if err := d.ReplaceText("hello", "x"); err != nil {
		t.Fatal(err)
	}

	r := p.Children[0].(*Run)
	if r.InstrText != "FORMTEXT hello" {
		t.Fatalf("instrText should not be changed, got %q", r.InstrText)
	}
}

func TestReplaceTextProcessesInstrTextWhenEnabled(t *testing.T) {
	const in = `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText xml:space="preserve"> FORMTEXT he</w:instrText></w:r>
<w:r><w:instrText xml:space="preserve">llo </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r>
<w:r><w:t>result</w:t></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r>
</w:p>`
	var p Paragraph
	if err := xml.Unmarshal(StringToBytes(in), &p); err != nil {
		t.Fatal(err)
	}
	d := New().WithDefaultTheme()
	d.Document.Body.Items = append(d.Document.Body.Items, &p)
	if err := d.ReplaceText("hello", "X", WithFieldCodeReplacement(true)); err != nil {
		t.Fatal(err)
	}
	pp := d.Document.Body.Items[0].(*Paragraph)
	gotInstr := paragraphInstrTexts(pp)
	if len(gotInstr) < 2 {
		t.Fatalf("expected at least 2 instrText runs, got %d", len(gotInstr))
	}
	if got := gotInstr[0]; got != " FORMTEXT X" {
		t.Fatalf("unexpected instrText in run 1: %q", got)
	}
	if got := gotInstr[1]; got != " " {
		t.Fatalf("unexpected instrText in run 2: %q", got)
	}
}

func TestReplaceTextSkipsNonWhitelistedFieldType(t *testing.T) {
	const in = `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText xml:space="preserve"> TOC hello </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r>
</w:p>`
	var p Paragraph
	if err := xml.Unmarshal(StringToBytes(in), &p); err != nil {
		t.Fatal(err)
	}
	d := New().WithDefaultTheme()
	d.Document.Body.Items = append(d.Document.Body.Items, &p)
	if err := d.ReplaceText("hello", "X", WithFieldCodeReplacement(true)); err != nil {
		t.Fatal(err)
	}
	pp := d.Document.Body.Items[0].(*Paragraph)
	gotInstr := paragraphInstrTexts(pp)
	if len(gotInstr) < 1 {
		t.Fatalf("expected instrText runs, got 0")
	}
	if got := gotInstr[0]; got != " TOC hello " {
		t.Fatalf("non-whitelisted field should not change, got %q", got)
	}
}

func TestReplaceTextSkipsMalformedFieldBoundaries(t *testing.T) {
	const in = `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText xml:space="preserve"> FORMTEXT hello </w:instrText></w:r>
</w:p>`
	var p Paragraph
	if err := xml.Unmarshal(StringToBytes(in), &p); err != nil {
		t.Fatal(err)
	}
	d := New().WithDefaultTheme()
	d.Document.Body.Items = append(d.Document.Body.Items, &p)
	if err := d.ReplaceText("hello", "X", WithFieldCodeReplacement(true)); err != nil {
		t.Fatal(err)
	}
	pp := d.Document.Body.Items[0].(*Paragraph)
	gotInstr := paragraphInstrTexts(pp)
	if len(gotInstr) < 1 {
		t.Fatalf("expected instrText runs, got 0")
	}
	if got := gotInstr[0]; got != " FORMTEXT hello " {
		t.Fatalf("malformed field should not change, got %q", got)
	}
}

func TestReplaceTextFieldCodeOptionsTakeEffect(t *testing.T) {
	const in = `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText xml:space="preserve"> MERGEFIELD Hello hello </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r>
</w:p>`
	var p Paragraph
	if err := xml.Unmarshal(StringToBytes(in), &p); err != nil {
		t.Fatal(err)
	}
	d := New().WithDefaultTheme()
	d.Document.Body.Items = append(d.Document.Body.Items, &p)
	if err := d.ReplaceText("hello", "X",
		WithFieldCodeReplacement(true),
		WithCaseSensitive(false),
		WithMaxReplacements(1),
	); err != nil {
		t.Fatal(err)
	}
	pp := d.Document.Body.Items[0].(*Paragraph)
	gotInstr := paragraphInstrTexts(pp)
	if len(gotInstr) < 1 {
		t.Fatalf("expected instrText runs, got 0")
	}
	if got := gotInstr[0]; got != " MERGEFIELD X hello " {
		t.Fatalf("field options not applied as expected, got %q", got)
	}
}

func paragraphInstrTexts(p *Paragraph) []string {
	items := p.Children
	if len(p.ordered) > 0 {
		items = p.ordered
	}
	out := make([]string, 0, 8)
	for _, it := range items {
		if r, ok := it.(*Run); ok && r.InstrText != "" {
			out = append(out, r.InstrText)
		}
	}
	return out
}

func TestReplaceTextRemovesEmptyRuns(t *testing.T) {
	d := New().WithDefaultTheme()
	p := d.AddParagraph()
	r1 := p.AddText("a")
	r2 := p.AddText("b")
	_ = r1

	if err := d.ReplaceText("ab", "X"); err != nil {
		t.Fatal(err)
	}
	if len(p.Children) != 1 {
		t.Fatalf("expected empty runs to be cleaned, got %d paragraph children", len(p.Children))
	}
	r, ok := p.Children[0].(*Run)
	if !ok || len(r.Children) != 1 || r.Children[0].(*Text).Text != "X" {
		t.Fatalf("unexpected remaining run content: %#v", p.Children[0])
	}
	if len(r2.Children) != 0 {
		t.Fatalf("expected second run to be emptied, got %d children", len(r2.Children))
	}
}

func TestReplaceTextPreservesStyleFromFirstMatchedRun(t *testing.T) {
	d := New().WithDefaultTheme()
	p := d.AddParagraph()
	r1 := p.AddText("he")
	r1.Color("FF0000")
	r2 := p.AddText("llo")
	r2.Color("00FF00")

	if err := d.ReplaceText("hello", "X"); err != nil {
		t.Fatal(err)
	}
	r, ok := p.Children[0].(*Run)
	if !ok {
		t.Fatalf("expected run, got %T", p.Children[0])
	}
	if r.RunProperties == nil || r.RunProperties.Color == nil || r.RunProperties.Color.Val != "FF0000" {
		t.Fatalf("replacement run should keep first matched run style, got %#v", r.RunProperties)
	}
}

func TestReplaceTextPreservesSpaceWithXMLSpace(t *testing.T) {
	d := New().WithDefaultTheme()
	p := d.AddParagraph()
	p.AddText("hello")

	if err := d.ReplaceText("hello", " hello "); err != nil {
		t.Fatal(err)
	}
	r := p.Children[0].(*Run)
	txt := r.Children[0].(*Text)
	if txt.XMLSpace != "preserve" {
		t.Fatalf("expected xml:space=preserve, got %q", txt.XMLSpace)
	}
}

func TestReplaceTextKeepsRawXMLNodesAfterRoundTrip(t *testing.T) {
	f, err := os.Open("testdata/roundtrip_unknown_nodes/sample_unknown.docx")
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()

	st, err := f.Stat()
	if err != nil {
		t.Fatal(err)
	}
	d, err := Parse(f, st.Size())
	if err != nil {
		t.Fatal(err)
	}
	if err := d.ReplaceText("roundtrip", "done"); err != nil {
		t.Fatal(err)
	}

	var out bytes.Buffer
	if _, err := d.WriteTo(&out); err != nil {
		t.Fatal(err)
	}
	d2, err := Parse(bytes.NewReader(out.Bytes()), int64(out.Len()))
	if err != nil {
		t.Fatal(err)
	}

	foundRaw := false
	for _, item := range d2.Document.Body.Items {
		if _, ok := item.(*RawXMLNode); ok {
			foundRaw = true
			break
		}
	}
	if !foundRaw {
		t.Fatal("expected preserved RawXMLNode in body items after replacement and round-trip")
	}
}
