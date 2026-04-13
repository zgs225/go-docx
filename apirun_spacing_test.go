package docx

import (
	"strings"
	"testing"
)

func TestRunSpacingWritesRunCharSpacingVal(t *testing.T) {
	doc := New().WithDefaultTheme()
	p := doc.AddParagraph()
	r := p.AddText("hello")

	if got := r.Spacing(240); got != r {
		t.Fatal("expected Run.Spacing to be chainable")
	}

	out, err := marshalXMLString(p)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(out, `<w:rPr><w:spacing w:val="240"></w:spacing></w:rPr>`) &&
		!strings.Contains(out, `<w:rPr><w:spacing w:val="240"/></w:rPr>`) {
		t.Fatalf("expected run spacing w:val in run properties, got: %s", out)
	}
	if strings.Contains(out, `<w:rPr><w:spacing w:line=`) {
		t.Fatalf("did not expect run spacing w:line in run properties, got: %s", out)
	}
}

func TestParagraphSpacingWritesParagraphLineSpacing(t *testing.T) {
	doc := New().WithDefaultTheme()
	p := doc.AddParagraph()

	if got := p.Spacing(360); got != p {
		t.Fatal("expected Paragraph.Spacing to be chainable")
	}
	p.AddText("hello")

	out, err := marshalXMLString(p)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(out, `<w:pPr><w:spacing w:line="360"></w:spacing></w:pPr>`) &&
		!strings.Contains(out, `<w:pPr><w:spacing w:line="360"/></w:pPr>`) {
		t.Fatalf("expected paragraph spacing w:line in paragraph properties, got: %s", out)
	}
	if strings.Contains(out, `<w:pPr><w:spacing w:val=`) {
		t.Fatalf("did not expect paragraph spacing w:val in paragraph properties, got: %s", out)
	}
}

func TestSpacingRunAndParagraphStaySeparated(t *testing.T) {
	doc := New().WithDefaultTheme()
	p := doc.AddParagraph().Spacing(480)
	p.AddText("hello").Spacing(120)

	out, err := marshalXMLString(p)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(out, `<w:pPr><w:spacing w:line="480"></w:spacing></w:pPr>`) &&
		!strings.Contains(out, `<w:pPr><w:spacing w:line="480"/></w:pPr>`) {
		t.Fatalf("expected paragraph spacing line in pPr, got: %s", out)
	}
	if !strings.Contains(out, `<w:rPr><w:spacing w:val="120"></w:spacing></w:rPr>`) &&
		!strings.Contains(out, `<w:rPr><w:spacing w:val="120"/></w:rPr>`) {
		t.Fatalf("expected run spacing val in rPr, got: %s", out)
	}
	if strings.Contains(out, `<w:rPr><w:spacing w:line=`) {
		t.Fatalf("did not expect run spacing line in rPr, got: %s", out)
	}
	if strings.Contains(out, `<w:pPr><w:spacing w:val=`) {
		t.Fatalf("did not expect paragraph spacing val in pPr, got: %s", out)
	}
}
