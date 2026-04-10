package docx

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"errors"
	"io"
	"regexp"
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

func TestSectPrTitlePgRoundTripAndUnknownOrder(t *testing.T) {
	input := `<w:sectPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:pgSz w:w="11906" w:h="16838"/><w:extA/><w:titlePg/><w:extB/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>`
	var s SectPr
	if err := xml.Unmarshal(StringToBytes(input), &s); err != nil {
		t.Fatal(err)
	}
	if s.TitlePg == nil {
		t.Fatal("expected titlePg to be parsed")
	}
	out, err := marshalXMLString(&s)
	if err != nil {
		t.Fatal(err)
	}
	got := directChildrenOf(strings.NewReader(out), "sectPr")
	want := []string{"pgSz", "extA", "titlePg", "extB", "pgMar"}
	if !equalStringSlice(got, want) {
		t.Fatalf("unexpected sectPr order with titlePg: got %v, want %v", got, want)
	}
}

func TestSettingsRoundTripAndUnknownOrder(t *testing.T) {
	input := `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:zoom w:percent="120"/><w:evenAndOddHeaders/><w:ext/></w:settings>`
	var s Settings
	if err := xml.Unmarshal(StringToBytes(input), &s); err != nil {
		t.Fatal(err)
	}
	if s.EvenAndOddHeaders == nil {
		t.Fatal("expected evenAndOddHeaders to be parsed")
	}
	out, err := marshalXMLString(&s)
	if err != nil {
		t.Fatal(err)
	}
	got := directChildrenOf(strings.NewReader(out), "settings")
	want := []string{"zoom", "evenAndOddHeaders", "ext"}
	if !equalStringSlice(got, want) {
		t.Fatalf("unexpected settings child order: got %v, want %v", got, want)
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

func TestAddPageNumberWritesPgNumTypeFormat(t *testing.T) {
	cases := []struct {
		name     string
		style    PageNumberStyle
		expected string
	}{
		{name: "arabic", style: PageNumberArabic, expected: "decimal"},
		{name: "roman", style: PageNumberRoman, expected: "lowerRoman"},
		{name: "ROMAN", style: PageNumberRomanUpper, expected: "upperRoman"},
		{name: "letter", style: PageNumberLetter, expected: "lowerLetter"},
		{name: "LETTER", style: PageNumberLetterUpper, expected: "upperLetter"},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			d := New().WithDefaultTheme().WithA4Page()
			if err := d.SetFooterText(FooterDefault, "Page: "); err != nil {
				t.Fatal(err)
			}
			if err := d.AddPageNumber(tc.style); err != nil {
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
			want := `w:pgNumType w:fmt="` + tc.expected + `"`
			if !strings.Contains(docXML, want) {
				t.Fatalf("missing pgNumType format %q in document.xml: %s", want, docXML)
			}
		})
	}
}

func TestHeaderFooterAlignmentAPI(t *testing.T) {
	d := New().WithDefaultTheme().WithA4Page()

	if err := d.SetHeaderAlignment(HeaderDefault, "center"); err != nil {
		t.Fatal(err)
	}
	if err := d.SetFooterAlignment(FooterDefault, "end"); err != nil {
		t.Fatal(err)
	}
	if err := d.AddPageNumberAligned(PageNumberArabic, "end"); err != nil {
		t.Fatal(err)
	}
	if err := d.SetHeaderAlignment(HeaderDefault, "left"); !errors.Is(err, errInvalidAlignment) {
		t.Fatalf("expected errInvalidAlignment, got: %v", err)
	}

	h, _ := d.GetHeader(HeaderDefault)
	if h == nil {
		t.Fatal("expected default header to be created")
	}
	ft, _ := d.GetFooter(FooterDefault)
	if ft == nil {
		t.Fatal("expected default footer to be created")
	}
	if h.Items == nil || len(h.Items) == 0 {
		t.Fatal("expected header paragraph to be auto-created")
	}
	if ft.Items == nil || len(ft.Items) == 0 {
		t.Fatal("expected footer paragraph to be auto-created")
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
	docXML := string(entries["word/document.xml"])
	if !strings.Contains(headerXML, `w:jc w:val="center"`) {
		t.Fatalf("expected centered header paragraph, got: %s", headerXML)
	}
	if !strings.Contains(footerXML, `w:jc w:val="end"`) {
		t.Fatalf("expected end-aligned footer paragraph, got: %s", footerXML)
	}
	if !strings.Contains(docXML, `w:pgNumType w:fmt="decimal"`) {
		t.Fatalf("expected decimal pgNumType, got: %s", docXML)
	}
}

func TestHeaderFooterRootAttrsNormalization(t *testing.T) {
	h := Header{
		attrs: []xml.Attr{
			{Name: xml.Name{Local: "_xmlns:w"}, Value: XMLNS_W},
			{Name: xml.Name{Space: "_xmlns", Local: "r"}, Value: XMLNS_R},
		},
	}
	h.AddText("h")
	out, err := marshalXMLString(&h)
	if err != nil {
		t.Fatal(err)
	}
	if strings.Contains(out, "_xmlns:") {
		t.Fatalf("unexpected invalid _xmlns attribute in header xml: %s", out)
	}
	if !strings.Contains(out, `xmlns:w="`+XMLNS_W+`"`) {
		t.Fatalf("missing normalized xmlns:w in header xml: %s", out)
	}
	if !strings.Contains(out, `xmlns:r="`+XMLNS_R+`"`) {
		t.Fatalf("missing normalized xmlns:r in header xml: %s", out)
	}

	ft := Footer{
		attrs: []xml.Attr{
			{Name: xml.Name{Local: "_xmlns:w"}, Value: XMLNS_W},
		},
	}
	ft.AddText("f")
	outFooter, err := marshalXMLString(&ft)
	if err != nil {
		t.Fatal(err)
	}
	if strings.Contains(outFooter, "_xmlns:") {
		t.Fatalf("unexpected invalid _xmlns attribute in footer xml: %s", outFooter)
	}
	if !strings.Contains(outFooter, `xmlns:w="`+XMLNS_W+`"`) {
		t.Fatalf("missing normalized xmlns:w in footer xml: %s", outFooter)
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

func TestParagraphSectPrRoundTrip(t *testing.T) {
	const customDoc = `<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:sectPr>
          <w:pgSz w:w="11906" w:h="16838"/>
          <w:customExt foo="bar"/>
        </w:sectPr>
      </w:pPr>
      <w:r><w:t>s1</w:t></w:r>
    </w:p>
    <w:p><w:r><w:t>s2</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
  </w:body>
</w:document>`
	data := buildDocxWithCustomDocumentXML(t, customDoc)
	doc, err := Parse(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatal(err)
	}
	if got := doc.SectionCount(); got != 2 {
		t.Fatalf("expected 2 sections, got %d", got)
	}

	var out bytes.Buffer
	if _, err := doc.WriteTo(&out); err != nil {
		t.Fatal(err)
	}
	entries, err := readZipEntriesFromBytes(out.Bytes())
	if err != nil {
		t.Fatal(err)
	}
	docXML := string(entries["word/document.xml"])
	if !strings.Contains(docXML, "<w:pPr>") || !strings.Contains(docXML, "<w:sectPr>") {
		t.Fatalf("expected paragraph sectPr to round-trip, got: %s", docXML)
	}
	if !strings.Contains(docXML, "customExt") {
		t.Fatalf("expected unknown node under paragraph sectPr to round-trip, got: %s", docXML)
	}
}

func TestSectionScopedHeaderFooterPageNumberAPI(t *testing.T) {
	const sectionedDoc = `<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:sectPr>
          <w:pgSz w:w="11906" w:h="16838"/>
        </w:sectPr>
      </w:pPr>
      <w:r><w:t>section-1</w:t></w:r>
    </w:p>
    <w:p><w:r><w:t>section-2</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
  </w:body>
</w:document>`
	data := buildDocxWithCustomDocumentXML(t, sectionedDoc)
	doc, err := Parse(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatal(err)
	}
	if got := doc.SectionCount(); got != 2 {
		t.Fatalf("expected 2 sections, got %d", got)
	}
	if err := doc.SetSectionHeaderText(0, HeaderDefault, "Header-S1"); err != nil {
		t.Fatal(err)
	}
	if err := doc.SetSectionFooterText(1, FooterDefault, "Footer-S2 "); err != nil {
		t.Fatal(err)
	}
	if err := doc.AddSectionPageNumberAligned(1, PageNumberArabic, "end"); err != nil {
		t.Fatal(err)
	}

	var out bytes.Buffer
	if _, err := doc.WriteTo(&out); err != nil {
		t.Fatal(err)
	}
	entries, err := readZipEntriesFromBytes(out.Bytes())
	if err != nil {
		t.Fatal(err)
	}
	docXML := string(entries["word/document.xml"])
	if strings.Count(docXML, "headerReference") < 1 || strings.Count(docXML, "footerReference") < 1 {
		t.Fatalf("expected section references in document.xml, got: %s", docXML)
	}
	if strings.Count(docXML, "<w:pgNumType ") != 1 {
		t.Fatalf("expected pgNumType on target section only, got: %s", docXML)
	}
	if !strings.Contains(docXML, `w:pgNumType w:fmt="decimal"`) {
		t.Fatalf("expected decimal pgNumType, got: %s", docXML)
	}

	doc2, err := Parse(bytes.NewReader(out.Bytes()), int64(out.Len()))
	if err != nil {
		t.Fatal(err)
	}
	h0, err := doc2.GetSectionHeader(0, HeaderDefault)
	if err != nil || h0 == nil {
		t.Fatalf("expected section 0 header, got h=%#v err=%v", h0, err)
	}
	hXML, err := marshalXMLString(h0)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(hXML, "Header-S1") {
		t.Fatalf("expected section 0 header content, got: %s", hXML)
	}

	f1, err := doc2.GetSectionFooter(1, FooterDefault)
	if err != nil || f1 == nil {
		t.Fatalf("expected section 1 footer, got f=%#v err=%v", f1, err)
	}
	fXML, err := marshalXMLString(f1)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(fXML, "Footer-S2") || !strings.Contains(fXML, `w:jc w:val="end"`) {
		t.Fatalf("expected section 1 footer content and alignment, got: %s", fXML)
	}
}

func TestSetSectionTitlePageOnlyAffectsTargetSection(t *testing.T) {
	const sectionedDoc = `<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:sectPr>
          <w:pgSz w:w="11906" w:h="16838"/>
        </w:sectPr>
      </w:pPr>
      <w:r><w:t>section-1</w:t></w:r>
    </w:p>
    <w:p><w:r><w:t>section-2</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
  </w:body>
</w:document>`
	data := buildDocxWithCustomDocumentXML(t, sectionedDoc)
	doc, err := Parse(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatal(err)
	}
	if err := doc.SetSectionTitlePage(1, true); err != nil {
		t.Fatal(err)
	}
	if err := doc.SetSectionTitlePage(3, true); !errors.Is(err, errSectionOutOfRange) {
		t.Fatalf("expected errSectionOutOfRange, got: %v", err)
	}

	var out bytes.Buffer
	if _, err := doc.WriteTo(&out); err != nil {
		t.Fatal(err)
	}
	doc2, err := Parse(bytes.NewReader(out.Bytes()), int64(out.Len()))
	if err != nil {
		t.Fatal(err)
	}
	sect0 := doc2.sectionByIndex(0)
	sect1 := doc2.sectionByIndex(1)
	if sect0 == nil || sect1 == nil {
		t.Fatalf("expected 2 sections after round-trip, got sect0=%v sect1=%v", sect0, sect1)
	}
	if sect0.TitlePg != nil {
		t.Fatal("expected section 0 titlePg unchanged")
	}
	if sect1.TitlePg == nil {
		t.Fatal("expected section 1 titlePg enabled")
	}
}

func TestSetEvenAndOddHeadersRequiresExistingSettingsPart(t *testing.T) {
	d := New().WithDefaultTheme().WithA4Page()
	if err := d.SetEvenAndOddHeaders(true); !errors.Is(err, errSettingsPartMissing) {
		t.Fatalf("expected errSettingsPartMissing, got: %v", err)
	}
}

func TestSetEvenAndOddHeadersToggleWithExistingSettings(t *testing.T) {
	data := buildDocxWithCustomParts(t, minimalSingleSectionDocumentXML(), map[string][]byte{
		"word/settings.xml": []byte(`<?xml version="1.0" encoding="UTF-8"?><w:settings xmlns:w="` + XMLNS_W + `"><w:zoom w:percent="120"/></w:settings>`),
	})
	doc, err := Parse(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatal(err)
	}
	if err := doc.SetEvenAndOddHeaders(true); err != nil {
		t.Fatal(err)
	}

	var outOn bytes.Buffer
	if _, err := doc.WriteTo(&outOn); err != nil {
		t.Fatal(err)
	}
	entriesOn, err := readZipEntriesFromBytes(outOn.Bytes())
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(string(entriesOn["word/settings.xml"]), "evenAndOddHeaders") {
		t.Fatalf("expected evenAndOddHeaders after enable, got: %s", string(entriesOn["word/settings.xml"]))
	}

	doc2, err := Parse(bytes.NewReader(outOn.Bytes()), int64(outOn.Len()))
	if err != nil {
		t.Fatal(err)
	}
	if err := doc2.SetEvenAndOddHeaders(false); err != nil {
		t.Fatal(err)
	}
	var outOff bytes.Buffer
	if _, err := doc2.WriteTo(&outOff); err != nil {
		t.Fatal(err)
	}
	entriesOff, err := readZipEntriesFromBytes(outOff.Bytes())
	if err != nil {
		t.Fatal(err)
	}
	if strings.Contains(string(entriesOff["word/settings.xml"]), "evenAndOddHeaders") {
		t.Fatalf("expected evenAndOddHeaders removed after disable, got: %s", string(entriesOff["word/settings.xml"]))
	}
}

func TestFirstEvenComboWithTitlePageAndSettingsSwitch(t *testing.T) {
	data := buildDocxWithCustomParts(t, sampleTwoSectionDocumentXML(), map[string][]byte{
		"word/settings.xml": []byte(`<?xml version="1.0" encoding="UTF-8"?><w:settings xmlns:w="` + XMLNS_W + `"/>`),
	})
	doc, err := Parse(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatal(err)
	}
	if err := doc.SetSectionHeaderText(0, HeaderDefault, "H0 Default"); err != nil {
		t.Fatal(err)
	}
	if err := doc.SetSectionHeaderText(0, HeaderFirst, "H0 First"); err != nil {
		t.Fatal(err)
	}
	if err := doc.SetSectionFooterText(1, FooterDefault, "F1 Default "); err != nil {
		t.Fatal(err)
	}
	if err := doc.SetSectionFooterText(1, FooterEven, "F1 Even "); err != nil {
		t.Fatal(err)
	}
	if err := doc.SetSectionTitlePage(0, true); err != nil {
		t.Fatal(err)
	}
	if err := doc.SetEvenAndOddHeaders(true); err != nil {
		t.Fatal(err)
	}
	if err := doc.AddSectionPageNumber(1, PageNumberRoman, FooterEven); err != nil {
		t.Fatal(err)
	}

	var out bytes.Buffer
	if _, err := doc.WriteTo(&out); err != nil {
		t.Fatal(err)
	}
	entries, err := readZipEntriesFromBytes(out.Bytes())
	if err != nil {
		t.Fatal(err)
	}
	docXML := string(entries["word/document.xml"])
	settingsXML := string(entries["word/settings.xml"])
	if !strings.Contains(docXML, "headerReference") || !strings.Contains(docXML, `w:type="first"`) {
		t.Fatalf("expected first headerReference in document.xml, got: %s", docXML)
	}
	if !strings.Contains(docXML, "footerReference") || !strings.Contains(docXML, `w:type="even"`) {
		t.Fatalf("expected even footerReference in document.xml, got: %s", docXML)
	}
	if !strings.Contains(docXML, "<w:titlePg") {
		t.Fatalf("expected titlePg in document.xml, got: %s", docXML)
	}
	if !strings.Contains(settingsXML, "evenAndOddHeaders") {
		t.Fatalf("expected evenAndOddHeaders in settings.xml, got: %s", settingsXML)
	}
	romanFound := false
	for name, data := range entries {
		if strings.HasPrefix(name, "word/footer") && strings.Contains(string(data), `PAGE \* roman`) {
			romanFound = true
			break
		}
	}
	if !romanFound {
		t.Fatal("expected roman page-number field in even footer part")
	}
}

func TestSectionFooterWriteIsolationWithCOW(t *testing.T) {
	data := buildDocxWithCustomDocumentXML(t, sampleTwoSectionDocumentXML())
	doc, err := Parse(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatal(err)
	}
	shared := doc.NewFooter()
	shared.AddText("Shared Footer")
	if err := doc.SetSectionFooter(0, FooterDefault, shared); err != nil {
		t.Fatal(err)
	}
	if err := doc.SetSectionFooter(1, FooterDefault, shared); err != nil {
		t.Fatal(err)
	}
	if err := doc.SetSectionFooterAlignment(0, FooterDefault, "end"); err != nil {
		t.Fatal(err)
	}

	f0, err := doc.GetSectionFooter(0, FooterDefault)
	if err != nil {
		t.Fatal(err)
	}
	f1, err := doc.GetSectionFooter(1, FooterDefault)
	if err != nil {
		t.Fatal(err)
	}
	if f0 == nil || f1 == nil {
		t.Fatalf("expected both section footers available, got f0=%v f1=%v", f0, f1)
	}
	if f0 == f1 {
		t.Fatal("expected copy-on-write clone, got shared footer pointer")
	}
	xml0, err := marshalXMLString(f0)
	if err != nil {
		t.Fatal(err)
	}
	xml1, err := marshalXMLString(f1)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(xml0, `w:jc w:val="end"`) {
		t.Fatalf("expected section 0 aligned footer, got: %s", xml0)
	}
	if strings.Contains(xml1, `w:jc w:val="end"`) {
		t.Fatalf("section 1 footer should stay unchanged, got: %s", xml1)
	}
}

func TestFooterPartDedupAcrossSections(t *testing.T) {
	data := buildDocxWithCustomDocumentXML(t, sampleTwoSectionDocumentXML())
	doc, err := Parse(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatal(err)
	}
	if err := doc.SetSectionFooterText(0, FooterDefault, "Same Footer"); err != nil {
		t.Fatal(err)
	}
	if err := doc.SetSectionFooterText(1, FooterDefault, "Same Footer"); err != nil {
		t.Fatal(err)
	}

	var out bytes.Buffer
	if _, err := doc.WriteTo(&out); err != nil {
		t.Fatal(err)
	}
	entries, err := readZipEntriesFromBytes(out.Bytes())
	if err != nil {
		t.Fatal(err)
	}
	footerParts := 0
	for name := range entries {
		if strings.HasPrefix(name, "word/footer") && strings.HasSuffix(strings.ToLower(name), ".xml") {
			footerParts++
		}
	}
	if footerParts != 1 {
		t.Fatalf("expected deduplicated single footer part, got %d", footerParts)
	}
	docXML := string(entries["word/document.xml"])
	relsXML := string(entries["word/_rels/document.xml.rels"])
	refIDRe := regexp.MustCompile(`<w:footerReference[^>]*w:type="default"[^>]*r:id="([^"]+)"`)
	matches := refIDRe.FindAllStringSubmatch(docXML, -1)
	if len(matches) < 2 {
		t.Fatalf("expected two default footer references for two sections, got: %s", docXML)
	}
	firstID := matches[0][1]
	for i := 1; i < len(matches); i++ {
		if matches[i][1] != firstID {
			t.Fatalf("expected all sections to reference one deduplicated RID, got %+v", matches)
		}
	}
	if strings.Count(relsXML, REL_FOOTER) != 1 {
		t.Fatalf("expected one footer relationship after dedup, got: %s", relsXML)
	}
}

func TestSectionAPIOutOfRange(t *testing.T) {
	d := New().WithDefaultTheme().WithA4Page()
	if d.SectionCount() != 1 {
		t.Fatalf("expected default single section")
	}
	if _, err := d.GetSectionHeader(1, HeaderDefault); !errors.Is(err, errSectionOutOfRange) {
		t.Fatalf("expected errSectionOutOfRange, got: %v", err)
	}
	if err := d.SetSectionFooterText(2, FooterDefault, "x"); !errors.Is(err, errSectionOutOfRange) {
		t.Fatalf("expected errSectionOutOfRange, got: %v", err)
	}
	if err := d.AddSectionPageNumber(3, PageNumberArabic); !errors.Is(err, errSectionOutOfRange) {
		t.Fatalf("expected errSectionOutOfRange, got: %v", err)
	}
}

func buildDocxWithCustomDocumentXML(t *testing.T, documentXML string) []byte {
	return buildDocxWithCustomParts(t, documentXML, nil)
}

func buildDocxWithCustomParts(t *testing.T, documentXML string, extra map[string][]byte) []byte {
	t.Helper()
	base := New().WithDefaultTheme().WithA4Page()
	base.AddParagraph().AddText("seed")

	var seed bytes.Buffer
	if _, err := base.WriteTo(&seed); err != nil {
		t.Fatal(err)
	}
	zr, err := zip.NewReader(bytes.NewReader(seed.Bytes()), int64(seed.Len()))
	if err != nil {
		t.Fatal(err)
	}

	entries := make(map[string][]byte, len(zr.File))
	for _, f := range zr.File {
		rc, err := f.Open()
		if err != nil {
			t.Fatal(err)
		}
		data, err := io.ReadAll(rc)
		_ = rc.Close()
		if err != nil {
			t.Fatal(err)
		}
		entries[f.Name] = data
	}
	entries["word/document.xml"] = []byte(documentXML)
	for name, data := range extra {
		entries[name] = data
	}

	var out bytes.Buffer
	zw := zip.NewWriter(&out)
	for name, data := range entries {
		w, err := zw.Create(name)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := w.Write(data); err != nil {
			t.Fatal(err)
		}
	}
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	return out.Bytes()
}

func minimalSingleSectionDocumentXML() string {
	return `<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>single</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
  </w:body>
</w:document>`
}

func sampleTwoSectionDocumentXML() string {
	return `<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:sectPr>
          <w:pgSz w:w="11906" w:h="16838"/>
        </w:sectPr>
      </w:pPr>
      <w:r><w:t>section-1</w:t></w:r>
    </w:p>
    <w:p><w:r><w:t>section-2</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
  </w:body>
</w:document>`
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
