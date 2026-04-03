package docx

import (
	"bytes"
	"encoding/xml"
	"strings"
	"testing"
)

func TestBodyPreserveUnknownOrder(t *testing.T) {
	input := `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>a</w:t></w:r></w:p><w:sdt><w:sdtContent><w:p><w:r><w:t>x</w:t></w:r></w:p></w:sdtContent></w:sdt><w:p><w:r><w:t>b</w:t></w:r></w:p></w:body></w:document>`

	doc := Document{
		XMLW:    XMLNS_W,
		XMLR:    XMLNS_R,
		XMLWP:   XMLNS_WP,
		XMLName: xml.Name{Space: XMLNS_W, Local: "document"},
	}
	if err := xml.Unmarshal(StringToBytes(input), &doc); err != nil {
		t.Fatal(err)
	}
	if len(doc.Body.Items) != 3 {
		t.Fatalf("expected 3 body items, got %d", len(doc.Body.Items))
	}
	if _, ok := doc.Body.Items[1].(*RawXMLNode); !ok {
		t.Fatalf("expected body[1] to be RawXMLNode, got %T", doc.Body.Items[1])
	}

	out, err := marshalXMLString(&doc)
	if err != nil {
		t.Fatal(err)
	}
	got := directChildrenOf(strings.NewReader(out), "body")
	want := []string{"p", "sdt", "p"}
	if !equalStringSlice(got, want) {
		t.Fatalf("unexpected body child order: got %v, want %v", got, want)
	}
}

func TestParagraphPreserveUnknownOrder(t *testing.T) {
	input := `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>a</w:t></w:r><w:proofErr w:type="spellStart"/><w:r><w:t>b</w:t></w:r></w:p>`
	var p Paragraph
	if err := xml.Unmarshal(StringToBytes(input), &p); err != nil {
		t.Fatal(err)
	}
	out, err := marshalXMLString(&p)
	if err != nil {
		t.Fatal(err)
	}
	got := directChildrenOf(strings.NewReader(out), "p")
	want := []string{"r", "proofErr", "r"}
	if !equalStringSlice(got, want) {
		t.Fatalf("unexpected p child order: got %v, want %v", got, want)
	}
}

func TestRunPreserveUnknownOrder(t *testing.T) {
	input := `<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:t>a</w:t><w:foo w:val="1"/><w:t>b</w:t></w:r>`
	var r Run
	if err := xml.Unmarshal(StringToBytes(input), &r); err != nil {
		t.Fatal(err)
	}
	out, err := marshalXMLString(&r)
	if err != nil {
		t.Fatal(err)
	}
	got := directChildrenOf(strings.NewReader(out), "r")
	want := []string{"t", "foo", "t"}
	if !equalStringSlice(got, want) {
		t.Fatalf("unexpected r child order: got %v, want %v", got, want)
	}
}

func TestTablePreserveUnknownOrder(t *testing.T) {
	input := `<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:tblPr/><w:tr><w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc></w:tr><w:unknown/><w:tr><w:tc><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`
	var tbl Table
	if err := xml.Unmarshal(StringToBytes(input), &tbl); err != nil {
		t.Fatal(err)
	}
	out, err := marshalXMLString(&tbl)
	if err != nil {
		t.Fatal(err)
	}
	got := directChildrenOf(strings.NewReader(out), "tbl")
	want := []string{"tblPr", "tr", "unknown", "tr"}
	if !equalStringSlice(got, want) {
		t.Fatalf("unexpected tbl child order: got %v, want %v", got, want)
	}
}

func TestTableRowPreserveUnknownOrder(t *testing.T) {
	input := `<w:tr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:trPr/><w:foo/><w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc></w:tr>`
	var tr WTableRow
	if err := xml.Unmarshal(StringToBytes(input), &tr); err != nil {
		t.Fatal(err)
	}
	out, err := marshalXMLString(&tr)
	if err != nil {
		t.Fatal(err)
	}
	got := directChildrenOf(strings.NewReader(out), "tr")
	want := []string{"trPr", "foo", "tc"}
	if !equalStringSlice(got, want) {
		t.Fatalf("unexpected tr child order: got %v, want %v", got, want)
	}
}

func TestTableCellPreserveUnknownOrder(t *testing.T) {
	input := `<w:tc xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:tcPr/><w:p><w:r><w:t>a</w:t></w:r></w:p><w:bar/><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc>`
	var tc WTableCell
	if err := xml.Unmarshal(StringToBytes(input), &tc); err != nil {
		t.Fatal(err)
	}
	out, err := marshalXMLString(&tc)
	if err != nil {
		t.Fatal(err)
	}
	got := directChildrenOf(strings.NewReader(out), "tc")
	want := []string{"tcPr", "p", "bar", "p"}
	if !equalStringSlice(got, want) {
		t.Fatalf("unexpected tc child order: got %v, want %v", got, want)
	}
}

func marshalXMLString(v interface{}) (string, error) {
	var buf bytes.Buffer
	if _, err := (marshaller{data: v}).WriteTo(&buf); err != nil {
		return "", err
	}
	return buf.String(), nil
}

func directChildrenOf(r *strings.Reader, parent string) []string {
	dec := xml.NewDecoder(r)
	depth := 0
	parentDepth := -1
	out := make([]string, 0, 8)
	for {
		tok, err := dec.Token()
		if err != nil {
			break
		}
		switch tt := tok.(type) {
		case xml.StartElement:
			depth++
			if tt.Name.Local == parent && parentDepth == -1 {
				parentDepth = depth
				continue
			}
			if parentDepth != -1 && depth == parentDepth+1 {
				out = append(out, tt.Name.Local)
			}
		case xml.EndElement:
			if parentDepth != -1 && tt.Name.Local == parent && depth == parentDepth {
				return out
			}
			depth--
		}
	}
	return out
}

func equalStringSlice(a, b []string) bool {
	if len(a) != len(b) {
		return false
	}
	for i := range a {
		if a[i] != b[i] {
			return false
		}
	}
	return true
}
