package docx

import (
	"bytes"
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
