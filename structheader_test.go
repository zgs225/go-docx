package docx

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"io"
	"strings"
	"testing"
)

func TestHeaderPreserveUnknownOrder(t *testing.T) {
	input := `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>a</w:t></w:r></w:p><w:custom/><w:p><w:r><w:t>b</w:t></w:r></w:p></w:hdr>`
	var h Header
	if err := xml.Unmarshal(StringToBytes(input), &h); err != nil {
		t.Fatal(err)
	}
	out, err := marshalXMLString(&h)
	if err != nil {
		t.Fatal(err)
	}
	got := directChildrenOf(strings.NewReader(out), "hdr")
	want := []string{"p", "custom", "p"}
	if !equalStringSlice(got, want) {
		t.Fatalf("unexpected header child order: got %v, want %v", got, want)
	}
}

func TestFooterPreserveUnknownOrder(t *testing.T) {
	input := `<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>a</w:t></w:r></w:p><w:custom/><w:p><w:r><w:t>b</w:t></w:r></w:p></w:ftr>`
	var ft Footer
	if err := xml.Unmarshal(StringToBytes(input), &ft); err != nil {
		t.Fatal(err)
	}
	out, err := marshalXMLString(&ft)
	if err != nil {
		t.Fatal(err)
	}
	got := directChildrenOf(strings.NewReader(out), "ftr")
	want := []string{"p", "custom", "p"}
	if !equalStringSlice(got, want) {
		t.Fatalf("unexpected footer child order: got %v, want %v", got, want)
	}
}

func TestSectPrHeaderFooterRefsRoundTrip(t *testing.T) {
	input := `<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:headerReference w:type="default" r:id="rId10"/><w:footerReference w:type="even" r:id="rId11"/><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>`
	var s SectPr
	if err := xml.Unmarshal(StringToBytes(input), &s); err != nil {
		t.Fatal(err)
	}
	if len(s.HeaderRefs) != 1 || s.HeaderRefs[0].RID != "rId10" || s.HeaderRefs[0].Type != "default" {
		t.Fatalf("unexpected header refs: %#v", s.HeaderRefs)
	}
	if len(s.FooterRefs) != 1 || s.FooterRefs[0].RID != "rId11" || s.FooterRefs[0].Type != "even" {
		t.Fatalf("unexpected footer refs: %#v", s.FooterRefs)
	}
	out, err := marshalXMLString(&s)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(out, "headerReference") || !strings.Contains(out, "footerReference") {
		t.Fatalf("expected header/footer refs in marshaled sectPr, got: %s", out)
	}
}

func TestHeaderFooterAPIAndPageNumberWrite(t *testing.T) {
	d := New().WithDefaultTheme().WithA4Page()
	if err := d.SetHeaderText(HeaderDefault, "Header Text"); err != nil {
		t.Fatal(err)
	}
	if err := d.SetFooterText(FooterDefault, "Page: "); err != nil {
		t.Fatal(err)
	}
	if err := d.AddPageNumber(PageNumberArabic); err != nil {
		t.Fatal(err)
	}

	h, err := d.GetHeader(HeaderDefault)
	if err != nil || h == nil {
		t.Fatalf("expected default header, got h=%#v err=%v", h, err)
	}
	ft, err := d.GetFooter(FooterDefault)
	if err != nil || ft == nil {
		t.Fatalf("expected default footer, got f=%#v err=%v", ft, err)
	}

	var buf bytes.Buffer
	if _, err := d.WriteTo(&buf); err != nil {
		t.Fatal(err)
	}
	entries, err := readZipEntriesFromBytes(buf.Bytes())
	if err != nil {
		t.Fatal(err)
	}

	headerXML := string(entries["word/header_default.xml"])
	footerXML := string(entries["word/footer_default.xml"])
	relsXML := string(entries["word/_rels/document.xml.rels"])
	docXML := string(entries["word/document.xml"])
	contentTypes := string(entries["[Content_Types].xml"])

	if !strings.Contains(headerXML, "Header Text") {
		t.Fatalf("header content not found: %s", headerXML)
	}
	if !strings.Contains(headerXML, `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`) {
		t.Fatalf("header namespace xmlns:w missing: %s", headerXML)
	}
	if !strings.Contains(footerXML, "Page:") {
		t.Fatalf("footer content not found: %s", footerXML)
	}
	if !strings.Contains(footerXML, `<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`) {
		t.Fatalf("footer namespace xmlns:w missing: %s", footerXML)
	}
	for _, needle := range []string{`fldCharType="begin"`, `fldCharType="separate"`, `fldCharType="end"`, "PAGE"} {
		if !strings.Contains(footerXML, needle) {
			t.Fatalf("footer page number field missing %q: %s", needle, footerXML)
		}
	}
	if !strings.Contains(relsXML, REL_HEADER) || !strings.Contains(relsXML, "header_default.xml") {
		t.Fatalf("header relationship missing: %s", relsXML)
	}
	if !strings.Contains(relsXML, REL_FOOTER) || !strings.Contains(relsXML, "footer_default.xml") {
		t.Fatalf("footer relationship missing: %s", relsXML)
	}
	if !strings.Contains(docXML, "headerReference") || !strings.Contains(docXML, "footerReference") {
		t.Fatalf("sectPr refs missing: %s", docXML)
	}
	if strings.Contains(docXML, "<PgSz") || !strings.Contains(docXML, "<w:pgSz") {
		t.Fatalf("sectPr page size tag should be w:pgSz, got: %s", docXML)
	}
	if !strings.Contains(contentTypes, "word/header_default.xml") || !strings.Contains(contentTypes, "word/footer_default.xml") {
		t.Fatalf("content types overrides missing: %s", contentTypes)
	}
	if strings.Contains(footerXML, `xmlns="w"`) || strings.Contains(footerXML, `xmlns:w="w"`) {
		t.Fatalf("footer fldChar namespace is invalid: %s", footerXML)
	}

	d2, err := Parse(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatal(err)
	}
	h2, _ := d2.GetHeader(HeaderDefault)
	ft2, _ := d2.GetFooter(FooterDefault)
	if h2 == nil || ft2 == nil {
		t.Fatalf("parsed doc should expose header/footer, got h=%#v f=%#v", h2, ft2)
	}
}

func TestAddPageNumberToSpecifiedFooterKind(t *testing.T) {
	d := New().WithDefaultTheme().WithA4Page()
	if err := d.SetFooterText(FooterEven, "E"); err != nil {
		t.Fatal(err)
	}
	if err := d.AddPageNumber(PageNumberRomanUpper, FooterEven); err != nil {
		t.Fatal(err)
	}

	var buf bytes.Buffer
	if _, err := d.WriteTo(&buf); err != nil {
		t.Fatal(err)
	}
	entries, err := readZipEntriesFromBytes(buf.Bytes())
	if err != nil {
		t.Fatal(err)
	}
	footerXML := string(entries["word/footer_even.xml"])
	if !strings.Contains(footerXML, `PAGE \* ROMAN`) {
		t.Fatalf("expected roman upper page field in even footer, got: %s", footerXML)
	}
}

func TestBodySectPrStaysAtTailAfterAddingContent(t *testing.T) {
	d := New().WithDefaultTheme().WithA4Page()
	_ = d.AddParagraph().AddText("p1")
	_ = d.AddTable(1, 1, 0, nil)
	if err := d.SetHeaderText(HeaderDefault, "h"); err != nil {
		t.Fatal(err)
	}
	if err := d.SetFooterText(FooterDefault, "f"); err != nil {
		t.Fatal(err)
	}

	var buf bytes.Buffer
	if _, err := d.WriteTo(&buf); err != nil {
		t.Fatal(err)
	}
	entries, err := readZipEntriesFromBytes(buf.Bytes())
	if err != nil {
		t.Fatal(err)
	}
	docXML := string(entries["word/document.xml"])

	bodyStart := strings.Index(docXML, "<w:body>")
	firstPara := strings.Index(docXML, "<w:p>")
	firstSect := strings.Index(docXML, "<w:sectPr>")
	bodyEnd := strings.Index(docXML, "</w:body>")
	lastSect := strings.LastIndex(docXML, "<w:sectPr>")
	if bodyStart < 0 || firstPara < 0 || firstSect < 0 || bodyEnd < 0 || lastSect < 0 {
		t.Fatalf("invalid document structure: %s", docXML)
	}
	if firstSect < firstPara {
		t.Fatalf("sectPr should not appear before paragraphs: %s", docXML)
	}
	if lastSect > bodyEnd {
		t.Fatalf("sectPr must be inside body: %s", docXML)
	}
	tail := docXML[lastSect:bodyEnd]
	if strings.Contains(tail, "<w:p>") || strings.Contains(tail, "<w:tbl>") {
		t.Fatalf("sectPr should be the final body-level element: %s", docXML)
	}
}

func readZipEntriesFromBytes(data []byte) (map[string][]byte, error) {
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return nil, err
	}
	out := make(map[string][]byte, len(zr.File))
	for _, f := range zr.File {
		rc, err := f.Open()
		if err != nil {
			return nil, err
		}
		b, err := io.ReadAll(rc)
		_ = rc.Close()
		if err != nil {
			return nil, err
		}
		out[f.Name] = b
	}
	return out, nil
}
