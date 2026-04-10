package main

import (
	"archive/zip"
	"bytes"
	"io"
	"os"
	"path/filepath"
	"regexp"
	"strings"
	"testing"

	"github.com/zgs225/go-docx"
)

func TestRoundTripCheckWithUnknownNodes(t *testing.T) {
	tmp := t.TempDir()
	inPath := filepath.Join(tmp, "input.docx")
	outPath := filepath.Join(tmp, "output.docx")

	if err := createSampleDocxWithUnknownNodes(inPath); err != nil {
		t.Fatal(err)
	}
	if err := roundTripOne(inPath, outPath); err != nil {
		t.Fatal(err)
	}
	issues, err := compareDocxStructure(inPath, outPath)
	if err != nil {
		t.Fatal(err)
	}
	if len(issues) != 0 {
		t.Fatalf("unexpected structural issues: %v", issues)
	}
}

func TestFieldCodeReplacementThenRoundTripCheck(t *testing.T) {
	tmp := t.TempDir()
	inPath := filepath.Join(tmp, "field-input.docx")
	replacedPath := filepath.Join(tmp, "field-replaced.docx")
	outPath := filepath.Join(tmp, "field-roundtrip.docx")

	if err := createSampleDocxWithFieldCodes(inPath); err != nil {
		t.Fatal(err)
	}

	inFile, err := os.Open(inPath)
	if err != nil {
		t.Fatal(err)
	}
	defer inFile.Close()
	st, err := inFile.Stat()
	if err != nil {
		t.Fatal(err)
	}
	doc, err := docx.Parse(inFile, st.Size())
	if err != nil {
		t.Fatal(err)
	}

	if err := doc.ReplaceText("hello", "X", docx.WithFieldCodeReplacement(true)); err != nil {
		t.Fatal(err)
	}

	outFile, err := os.Create(replacedPath)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := doc.WriteTo(outFile); err != nil {
		_ = outFile.Close()
		t.Fatal(err)
	}
	_ = outFile.Close()

	replacedEntries, err := readZipEntries(replacedPath)
	if err != nil {
		t.Fatal(err)
	}
	replacedXML := string(replacedEntries["word/document.xml"])
	if !strings.Contains(replacedXML, "MERGEFIELD X") {
		t.Fatalf("expected replaced field code in document.xml, got: %s", replacedXML)
	}
	if !strings.Contains(replacedXML, `fldCharType="begin"`) || !strings.Contains(replacedXML, `fldCharType="end"`) {
		t.Fatal("expected field boundary markers to be preserved after replacement")
	}

	if err := roundTripOne(replacedPath, outPath); err != nil {
		t.Fatal(err)
	}
	issues, err := compareDocxStructure(replacedPath, outPath)
	if err != nil {
		t.Fatal(err)
	}
	if len(issues) != 0 {
		t.Fatalf("unexpected structural issues after field replacement and round-trip: %v", issues)
	}
}

func TestTableSpanBorderThenRoundTripCheck(t *testing.T) {
	tmp := t.TempDir()
	inPath := filepath.Join(tmp, "table-input.docx")
	replacedPath := filepath.Join(tmp, "table-updated.docx")
	outPath := filepath.Join(tmp, "table-roundtrip.docx")

	if err := createSampleDocxWithUnknownNodes(inPath); err != nil {
		t.Fatal(err)
	}

	inFile, err := os.Open(inPath)
	if err != nil {
		t.Fatal(err)
	}
	defer inFile.Close()

	st, err := inFile.Stat()
	if err != nil {
		t.Fatal(err)
	}
	doc, err := docx.Parse(inFile, st.Size())
	if err != nil {
		t.Fatal(err)
	}

	var target *docx.WTableCell
	for _, item := range doc.Document.Body.Items {
		tbl, ok := item.(*docx.Table)
		if !ok || len(tbl.TableRows) == 0 || len(tbl.TableRows[0].TableCells) == 0 {
			continue
		}
		target = tbl.TableRows[0].TableCells[0]
		break
	}
	if target == nil {
		t.Fatal("expected at least one table cell in sample")
	}

	target.SetColSpan(2).
		SetRowSpanRestart().
		Padding(120, 180, 240, 300).
		SetCellBordersSame("single", 8, 0, "000000")

	outFile, err := os.Create(replacedPath)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := doc.WriteTo(outFile); err != nil {
		_ = outFile.Close()
		t.Fatal(err)
	}
	_ = outFile.Close()

	replacedEntries, err := readZipEntries(replacedPath)
	if err != nil {
		t.Fatal(err)
	}
	replacedXML := string(replacedEntries["word/document.xml"])
	if !strings.Contains(replacedXML, `gridSpan`) {
		t.Fatalf("expected gridSpan in document.xml, got: %s", replacedXML)
	}
	if !strings.Contains(replacedXML, `vMerge`) || !strings.Contains(replacedXML, `restart`) {
		t.Fatalf("expected vMerge restart in document.xml, got: %s", replacedXML)
	}
	if !strings.Contains(replacedXML, `tcBorders`) {
		t.Fatalf("expected tcBorders in document.xml, got: %s", replacedXML)
	}
	if !strings.Contains(replacedXML, `extNode`) {
		t.Fatalf("expected unknown node preserved in document.xml, got: %s", replacedXML)
	}

	if err := roundTripOne(replacedPath, outPath); err != nil {
		t.Fatal(err)
	}
	issues, err := compareDocxStructure(replacedPath, outPath)
	if err != nil {
		t.Fatal(err)
	}
	if len(issues) != 0 {
		t.Fatalf("unexpected structural issues after table span/border update and round-trip: %v", issues)
	}
}

func TestTableSpanBorderClearThenRoundTripCheck(t *testing.T) {
	tmp := t.TempDir()
	inPath := filepath.Join(tmp, "table-input-clear.docx")
	replacedPath := filepath.Join(tmp, "table-updated-clear.docx")
	outPath := filepath.Join(tmp, "table-roundtrip-clear.docx")

	if err := createSampleDocxWithUnknownNodes(inPath); err != nil {
		t.Fatal(err)
	}

	inFile, err := os.Open(inPath)
	if err != nil {
		t.Fatal(err)
	}
	defer inFile.Close()

	st, err := inFile.Stat()
	if err != nil {
		t.Fatal(err)
	}
	doc, err := docx.Parse(inFile, st.Size())
	if err != nil {
		t.Fatal(err)
	}

	var target *docx.WTableCell
	for _, item := range doc.Document.Body.Items {
		tbl, ok := item.(*docx.Table)
		if !ok || len(tbl.TableRows) == 0 || len(tbl.TableRows[0].TableCells) == 0 {
			continue
		}
		target = tbl.TableRows[0].TableCells[0]
		break
	}
	if target == nil {
		t.Fatal("expected at least one table cell in sample")
	}

	target.SetColSpan(2).
		SetRowSpanContinue().
		SetCellBordersSame("single", 8, 0, "000000").
		ClearCellBorders()

	outFile, err := os.Create(replacedPath)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := doc.WriteTo(outFile); err != nil {
		_ = outFile.Close()
		t.Fatal(err)
	}
	_ = outFile.Close()

	replacedEntries, err := readZipEntries(replacedPath)
	if err != nil {
		t.Fatal(err)
	}
	replacedXML := string(replacedEntries["word/document.xml"])
	if !strings.Contains(replacedXML, `gridSpan`) {
		t.Fatalf("expected gridSpan in document.xml, got: %s", replacedXML)
	}
	if !strings.Contains(replacedXML, `vMerge`) {
		t.Fatalf("expected vMerge in document.xml, got: %s", replacedXML)
	}
	if strings.Contains(replacedXML, `tcBorders`) {
		t.Fatalf("expected tcBorders cleared in document.xml, got: %s", replacedXML)
	}
	if !strings.Contains(replacedXML, `extNode`) {
		t.Fatalf("expected unknown node preserved in document.xml, got: %s", replacedXML)
	}

	if err := roundTripOne(replacedPath, outPath); err != nil {
		t.Fatal(err)
	}
	issues, err := compareDocxStructure(replacedPath, outPath)
	if err != nil {
		t.Fatal(err)
	}
	if len(issues) != 0 {
		t.Fatalf("unexpected structural issues after table span/border clear and round-trip: %v", issues)
	}
}

func TestTableDefaultPaddingAndLayoutThenRoundTripCheck(t *testing.T) {
	tmp := t.TempDir()
	inPath := filepath.Join(tmp, "table-input-layout.docx")
	replacedPath := filepath.Join(tmp, "table-updated-layout.docx")
	outPath := filepath.Join(tmp, "table-roundtrip-layout.docx")

	if err := createSampleDocxWithUnknownNodes(inPath); err != nil {
		t.Fatal(err)
	}

	inFile, err := os.Open(inPath)
	if err != nil {
		t.Fatal(err)
	}
	defer inFile.Close()

	st, err := inFile.Stat()
	if err != nil {
		t.Fatal(err)
	}
	doc, err := docx.Parse(inFile, st.Size())
	if err != nil {
		t.Fatal(err)
	}

	var target *docx.Table
	for _, item := range doc.Document.Body.Items {
		tbl, ok := item.(*docx.Table)
		if !ok {
			continue
		}
		target = tbl
		break
	}
	if target == nil {
		t.Fatal("expected at least one table in sample")
	}

	target.SetDefaultCellPadding(120, 180, 240, 300).
		SetLayoutFixed().
		SetWidthTwips(7200)
	target.TableRows[0].TableCells[0].Padding(10, 20, 30, 40)

	outFile, err := os.Create(replacedPath)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := doc.WriteTo(outFile); err != nil {
		_ = outFile.Close()
		t.Fatal(err)
	}
	_ = outFile.Close()

	replacedEntries, err := readZipEntries(replacedPath)
	if err != nil {
		t.Fatal(err)
	}
	replacedXML := string(replacedEntries["word/document.xml"])
	if !strings.Contains(replacedXML, `tblCellMar`) {
		t.Fatalf("expected tblCellMar in document.xml, got: %s", replacedXML)
	}
	if !strings.Contains(replacedXML, `tblLayout`) || !strings.Contains(replacedXML, `fixed`) {
		t.Fatalf("expected fixed tblLayout in document.xml, got: %s", replacedXML)
	}
	if !strings.Contains(replacedXML, `tcMar`) {
		t.Fatalf("expected cell-level tcMar in document.xml, got: %s", replacedXML)
	}
	if !strings.Contains(replacedXML, `extNode`) {
		t.Fatalf("expected unknown node preserved in document.xml, got: %s", replacedXML)
	}

	if err := roundTripOne(replacedPath, outPath); err != nil {
		t.Fatal(err)
	}
	issues, err := compareDocxStructure(replacedPath, outPath)
	if err != nil {
		t.Fatal(err)
	}
	if len(issues) != 0 {
		t.Fatalf("unexpected structural issues after table default padding/layout update and round-trip: %v", issues)
	}
}

func TestTableRepeatHeaderAndUnifiedLayoutThenRoundTripCheck(t *testing.T) {
	tmp := t.TempDir()
	inPath := filepath.Join(tmp, "table-input-repeat-header.docx")
	replacedPath := filepath.Join(tmp, "table-updated-repeat-header.docx")
	outPath := filepath.Join(tmp, "table-roundtrip-repeat-header.docx")

	if err := createSampleDocxWithUnknownNodes(inPath); err != nil {
		t.Fatal(err)
	}

	inFile, err := os.Open(inPath)
	if err != nil {
		t.Fatal(err)
	}
	defer inFile.Close()

	st, err := inFile.Stat()
	if err != nil {
		t.Fatal(err)
	}
	doc, err := docx.Parse(inFile, st.Size())
	if err != nil {
		t.Fatal(err)
	}

	var target *docx.Table
	for _, item := range doc.Document.Body.Items {
		tbl, ok := item.(*docx.Table)
		if !ok {
			continue
		}
		target = tbl
		break
	}
	if target == nil {
		t.Fatal("expected at least one table in sample")
	}

	target.SetLayout(docx.TableLayoutOptions{
		Mode:          docx.TableLayoutModeFixed,
		WidthTwips:    7200,
		Justification: "center",
		DefaultCellPadding: &docx.TablePadding{
			Top: 120, Right: 180, Bottom: 240, Left: 300,
		},
	})
	target.SetRepeatHeader(0, true)
	target.SetRowLayout(0, docx.RowLayoutOptions{
		Justification: "center",
		HeightTwips:   900,
		HeightRule:    "exact",
	})

	outFile, err := os.Create(replacedPath)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := doc.WriteTo(outFile); err != nil {
		_ = outFile.Close()
		t.Fatal(err)
	}
	_ = outFile.Close()

	replacedEntries, err := readZipEntries(replacedPath)
	if err != nil {
		t.Fatal(err)
	}
	replacedXML := string(replacedEntries["word/document.xml"])
	for _, needle := range []string{`tblHeader`, `trHeight`, `tblLayout`, `tblCellMar`} {
		if !strings.Contains(replacedXML, needle) {
			t.Fatalf("expected %q in document.xml, got: %s", needle, replacedXML)
		}
	}
	if !strings.Contains(replacedXML, `extNode`) {
		t.Fatalf("expected unknown node preserved in document.xml, got: %s", replacedXML)
	}

	if err := roundTripOne(replacedPath, outPath); err != nil {
		t.Fatal(err)
	}
	issues, err := compareDocxStructure(replacedPath, outPath)
	if err != nil {
		t.Fatal(err)
	}
	if len(issues) != 0 {
		t.Fatalf("unexpected structural issues after repeat-header/layout update and round-trip: %v", issues)
	}
}

func TestTableRepeatHeaderAndUnifiedLayoutClearSwitchThenRoundTripCheck(t *testing.T) {
	tmp := t.TempDir()
	inPath := filepath.Join(tmp, "table-input-repeat-header-clear.docx")
	replacedPath := filepath.Join(tmp, "table-updated-repeat-header-clear.docx")
	outPath := filepath.Join(tmp, "table-roundtrip-repeat-header-clear.docx")

	if err := createSampleDocxWithUnknownNodes(inPath); err != nil {
		t.Fatal(err)
	}

	inFile, err := os.Open(inPath)
	if err != nil {
		t.Fatal(err)
	}
	defer inFile.Close()

	st, err := inFile.Stat()
	if err != nil {
		t.Fatal(err)
	}
	doc, err := docx.Parse(inFile, st.Size())
	if err != nil {
		t.Fatal(err)
	}

	var target *docx.Table
	for _, item := range doc.Document.Body.Items {
		tbl, ok := item.(*docx.Table)
		if !ok {
			continue
		}
		target = tbl
		break
	}
	if target == nil {
		t.Fatal("expected at least one table in sample")
	}

	target.SetLayout(docx.TableLayoutOptions{
		Mode:          docx.TableLayoutModeFixed,
		WidthTwips:    7200,
		Justification: "center",
		DefaultCellPadding: &docx.TablePadding{
			Top: 120, Right: 180, Bottom: 240, Left: 300,
		},
	})
	target.SetRepeatHeader(0, true)
	target.SetRowLayout(0, docx.RowLayoutOptions{
		Justification: "center",
		HeightTwips:   900,
		HeightRule:    "exact",
	})
	repeatFalse := false
	target.SetRowLayout(0, docx.RowLayoutOptions{
		HeightTwips:  0,
		RepeatHeader: &repeatFalse,
	})
	target.SetLayout(docx.TableLayoutOptions{
		Mode:       docx.TableLayoutModeAutofit,
		WidthTwips: 0,
	})

	outFile, err := os.Create(replacedPath)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := doc.WriteTo(outFile); err != nil {
		_ = outFile.Close()
		t.Fatal(err)
	}
	_ = outFile.Close()

	replacedEntries, err := readZipEntries(replacedPath)
	if err != nil {
		t.Fatal(err)
	}
	replacedXML := string(replacedEntries["word/document.xml"])
	if strings.Contains(replacedXML, `tblHeader`) {
		t.Fatalf("expected tblHeader cleared in document.xml, got: %s", replacedXML)
	}
	if strings.Contains(replacedXML, `trHeight`) {
		t.Fatalf("expected trHeight cleared in document.xml, got: %s", replacedXML)
	}
	if !strings.Contains(replacedXML, `tblLayout`) || !strings.Contains(replacedXML, `autofit`) {
		t.Fatalf("expected autofit tblLayout in document.xml, got: %s", replacedXML)
	}
	if !strings.Contains(replacedXML, `tblCellMar`) {
		t.Fatalf("expected tblCellMar preserved after layout switch, got: %s", replacedXML)
	}
	if !strings.Contains(replacedXML, `extNode`) {
		t.Fatalf("expected unknown node preserved in document.xml, got: %s", replacedXML)
	}

	if err := roundTripOne(replacedPath, outPath); err != nil {
		t.Fatal(err)
	}
	issues, err := compareDocxStructure(replacedPath, outPath)
	if err != nil {
		t.Fatal(err)
	}
	if len(issues) != 0 {
		t.Fatalf("unexpected structural issues after repeat-header/layout clear-switch and round-trip: %v", issues)
	}
}

func TestHeaderFooterPageNumberThenRoundTripCheck(t *testing.T) {
	tmp := t.TempDir()
	inPath := filepath.Join(tmp, "hf-input.docx")
	replacedPath := filepath.Join(tmp, "hf-updated.docx")
	outPath := filepath.Join(tmp, "hf-roundtrip.docx")

	if err := createSampleDocxWithHeaderFooter(inPath); err != nil {
		t.Fatal(err)
	}

	inFile, err := os.Open(inPath)
	if err != nil {
		t.Fatal(err)
	}
	defer inFile.Close()

	st, err := inFile.Stat()
	if err != nil {
		t.Fatal(err)
	}
	doc, err := docx.Parse(inFile, st.Size())
	if err != nil {
		t.Fatal(err)
	}

	if err := doc.SetHeaderText(docx.HeaderDefault, "Updated Header"); err != nil {
		t.Fatal(err)
	}
	if err := doc.SetHeaderAlignment(docx.HeaderDefault, "center"); err != nil {
		t.Fatal(err)
	}
	if err := doc.SetFooterText(docx.FooterDefault, "Page: "); err != nil {
		t.Fatal(err)
	}
	if err := doc.AddPageNumberAligned(docx.PageNumberArabic, "end"); err != nil {
		t.Fatal(err)
	}

	outFile, err := os.Create(replacedPath)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := doc.WriteTo(outFile); err != nil {
		_ = outFile.Close()
		t.Fatal(err)
	}
	_ = outFile.Close()

	entries, err := readZipEntries(replacedPath)
	if err != nil {
		t.Fatal(err)
	}
	headerXML := string(entries["word/header_default.xml"])
	footerXML := string(entries["word/footer_default.xml"])
	docXML := string(entries["word/document.xml"])
	if !strings.Contains(headerXML, "Updated Header") {
		t.Fatalf("updated header text missing: %s", headerXML)
	}
	if !strings.Contains(headerXML, `w:jc w:val="center"`) {
		t.Fatalf("header alignment missing: %s", headerXML)
	}
	if strings.Contains(headerXML, `_xmlns:`) {
		t.Fatalf("unexpected invalid namespace attr in header xml: %s", headerXML)
	}
	for _, needle := range []string{`fldCharType="begin"`, `fldCharType="separate"`, `fldCharType="end"`, "PAGE"} {
		if !strings.Contains(footerXML, needle) {
			t.Fatalf("footer page number field missing %q: %s", needle, footerXML)
		}
	}
	if !strings.Contains(footerXML, `w:jc w:val="end"`) {
		t.Fatalf("footer alignment missing: %s", footerXML)
	}
	if strings.Contains(footerXML, `_xmlns:`) {
		t.Fatalf("unexpected invalid namespace attr in footer xml: %s", footerXML)
	}
	if !strings.Contains(docXML, "headerReference") || !strings.Contains(docXML, "footerReference") {
		t.Fatalf("document sectPr refs missing: %s", docXML)
	}
	if !strings.Contains(docXML, `w:pgNumType w:fmt="decimal"`) {
		t.Fatalf("document pgNumType missing: %s", docXML)
	}

	if err := roundTripOne(replacedPath, outPath); err != nil {
		t.Fatal(err)
	}
	issues, err := compareDocxStructure(replacedPath, outPath)
	if err != nil {
		t.Fatal(err)
	}
	if len(issues) != 0 {
		t.Fatalf("unexpected structural issues after header/footer/page number update and round-trip: %v", issues)
	}
}

func TestMultiSectionHeaderFooterPageNumberThenRoundTripCheck(t *testing.T) {
	tmp := t.TempDir()
	inPath := filepath.Join(tmp, "multisec-input.docx")
	updatedPath := filepath.Join(tmp, "multisec-updated.docx")
	outPath := filepath.Join(tmp, "multisec-roundtrip.docx")

	if err := createSampleDocxWithSections(inPath); err != nil {
		t.Fatal(err)
	}

	inFile, err := os.Open(inPath)
	if err != nil {
		t.Fatal(err)
	}
	defer inFile.Close()
	st, err := inFile.Stat()
	if err != nil {
		t.Fatal(err)
	}
	doc, err := docx.Parse(inFile, st.Size())
	if err != nil {
		t.Fatal(err)
	}
	if doc.SectionCount() != 2 {
		t.Fatalf("expected 2 sections, got %d", doc.SectionCount())
	}
	if err := doc.SetSectionHeaderText(0, docx.HeaderDefault, "S1 Header"); err != nil {
		t.Fatal(err)
	}
	if err := doc.SetSectionFooterText(1, docx.FooterDefault, "S2 Page "); err != nil {
		t.Fatal(err)
	}
	if err := doc.AddSectionPageNumberAligned(1, docx.PageNumberArabic, "end"); err != nil {
		t.Fatal(err)
	}

	outFile, err := os.Create(updatedPath)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := doc.WriteTo(outFile); err != nil {
		_ = outFile.Close()
		t.Fatal(err)
	}
	_ = outFile.Close()

	entries, err := readZipEntries(updatedPath)
	if err != nil {
		t.Fatal(err)
	}
	docXML := string(entries["word/document.xml"])
	if strings.Count(docXML, "<w:sectPr") < 2 {
		t.Fatalf("expected 2 sectPr nodes in document.xml, got: %s", docXML)
	}
	if !strings.Contains(docXML, `w:pgNumType w:fmt="decimal"`) {
		t.Fatalf("expected section page number format in document.xml, got: %s", docXML)
	}
	headerFound := false
	footerFound := false
	for name, data := range entries {
		if strings.HasPrefix(name, "word/header") && strings.Contains(string(data), "S1 Header") {
			headerFound = true
		}
		if strings.HasPrefix(name, "word/footer") && strings.Contains(string(data), "S2 Page") {
			footerFound = true
		}
	}
	if !headerFound || !footerFound {
		t.Fatalf("expected updated multi-section header/footer parts, headerFound=%v footerFound=%v", headerFound, footerFound)
	}

	if err := roundTripOne(updatedPath, outPath); err != nil {
		t.Fatal(err)
	}
	issues, err := compareDocxStructure(updatedPath, outPath)
	if err != nil {
		t.Fatal(err)
	}
	if len(issues) != 0 {
		t.Fatalf("unexpected structural issues after multi-section update and round-trip: %v", issues)
	}
}

func TestMultiSectionFirstEvenCowDedupThenRoundTripCheck(t *testing.T) {
	tmp := t.TempDir()
	inPath := filepath.Join(tmp, "multisec-feature-input.docx")
	updatedPath := filepath.Join(tmp, "multisec-feature-updated.docx")
	outPath := filepath.Join(tmp, "multisec-feature-roundtrip.docx")

	if err := createSampleDocxWithSectionsAndSettings(inPath); err != nil {
		t.Fatal(err)
	}

	inFile, err := os.Open(inPath)
	if err != nil {
		t.Fatal(err)
	}
	defer inFile.Close()
	st, err := inFile.Stat()
	if err != nil {
		t.Fatal(err)
	}
	doc, err := docx.Parse(inFile, st.Size())
	if err != nil {
		t.Fatal(err)
	}

	if doc.SectionCount() != 2 {
		t.Fatalf("expected 2 sections, got %d", doc.SectionCount())
	}
	shared := doc.NewFooter()
	shared.AddText("Shared")
	if err := doc.SetSectionFooter(0, docx.FooterDefault, shared); err != nil {
		t.Fatal(err)
	}
	if err := doc.SetSectionFooter(1, docx.FooterDefault, shared); err != nil {
		t.Fatal(err)
	}
	if err := doc.SetSectionFooterAlignment(0, docx.FooterDefault, "end"); err != nil {
		t.Fatal(err)
	}
	if err := doc.SetSectionTitlePage(0, true); err != nil {
		t.Fatal(err)
	}
	if err := doc.SetEvenAndOddHeaders(true); err != nil {
		t.Fatal(err)
	}
	if err := doc.SetSectionHeaderText(0, docx.HeaderFirst, "S1 First Header"); err != nil {
		t.Fatal(err)
	}
	if err := doc.SetSectionFooterText(1, docx.FooterEven, "S2 Even "); err != nil {
		t.Fatal(err)
	}
	if err := doc.AddSectionPageNumber(1, docx.PageNumberRoman, docx.FooterEven); err != nil {
		t.Fatal(err)
	}
	if err := doc.SetSectionFooterText(0, docx.FooterDefault, "Same Footer"); err != nil {
		t.Fatal(err)
	}
	if err := doc.SetSectionFooterText(1, docx.FooterDefault, "Same Footer"); err != nil {
		t.Fatal(err)
	}

	outFile, err := os.Create(updatedPath)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := doc.WriteTo(outFile); err != nil {
		_ = outFile.Close()
		t.Fatal(err)
	}
	_ = outFile.Close()

	entries, err := readZipEntries(updatedPath)
	if err != nil {
		t.Fatal(err)
	}
	docXML := string(entries["word/document.xml"])
	settingsXML := string(entries["word/settings.xml"])
	if !strings.Contains(docXML, "<w:titlePg") {
		t.Fatalf("expected titlePg in document.xml, got: %s", docXML)
	}
	if !strings.Contains(docXML, "headerReference") || !strings.Contains(docXML, `w:type="first"`) {
		t.Fatalf("expected first header reference in document.xml, got: %s", docXML)
	}
	if !strings.Contains(docXML, "footerReference") || !strings.Contains(docXML, `w:type="even"`) {
		t.Fatalf("expected even footer reference in document.xml, got: %s", docXML)
	}
	if !strings.Contains(settingsXML, "evenAndOddHeaders") {
		t.Fatalf("expected evenAndOddHeaders in settings.xml, got: %s", settingsXML)
	}
	footerParts := 0
	for name := range entries {
		if strings.HasPrefix(name, "word/footer") && strings.HasSuffix(strings.ToLower(name), ".xml") {
			footerParts++
		}
	}
	if footerParts != 2 {
		t.Fatalf("expected deduped default footer + even footer => 2 parts, got %d", footerParts)
	}
	refIDRe := regexp.MustCompile(`<w:footerReference[^>]*w:type="default"[^>]*r:id="([^"]+)"`)
	matches := refIDRe.FindAllStringSubmatch(docXML, -1)
	if len(matches) < 2 {
		t.Fatalf("expected two default footer references, got: %s", docXML)
	}
	if matches[0][1] != matches[1][1] {
		t.Fatalf("expected deduped default footer RID, got %+v", matches)
	}

	if err := roundTripOne(updatedPath, outPath); err != nil {
		t.Fatal(err)
	}
	issues, err := compareDocxStructure(updatedPath, outPath)
	if err != nil {
		t.Fatal(err)
	}
	if len(issues) != 0 {
		t.Fatalf("unexpected structural issues after multi-section first/even+cow+dedup update and round-trip: %v", issues)
	}
}

func TestCombinedReplaceTableHeaderFooterThenRoundTripCheck(t *testing.T) {
	tmp := t.TempDir()
	inPath := filepath.Join(tmp, "combo-input.docx")
	updatedPath := filepath.Join(tmp, "combo-updated.docx")
	outPath := filepath.Join(tmp, "combo-roundtrip.docx")

	if err := createSampleDocxWithUnknownNodes(inPath); err != nil {
		t.Fatal(err)
	}

	inFile, err := os.Open(inPath)
	if err != nil {
		t.Fatal(err)
	}
	defer inFile.Close()

	st, err := inFile.Stat()
	if err != nil {
		t.Fatal(err)
	}
	doc, err := docx.Parse(inFile, st.Size())
	if err != nil {
		t.Fatal(err)
	}

	doc.AddParagraph().AddText("Hello {{name}}")
	if err := doc.ReplacePlaceholder(map[string]string{"name": "Codex"}); err != nil {
		t.Fatal(err)
	}

	var target *docx.WTableCell
	for _, item := range doc.Document.Body.Items {
		tbl, ok := item.(*docx.Table)
		if !ok || len(tbl.TableRows) == 0 || len(tbl.TableRows[0].TableCells) == 0 {
			continue
		}
		target = tbl.TableRows[0].TableCells[0]
		break
	}
	if target == nil {
		t.Fatal("expected at least one table cell in sample")
	}
	target.SetColSpan(2).
		SetRowSpanRestart().
		Padding(30, 40, 50, 60).
		SetCellBordersSame("single", 8, 0, "000000")

	if err := doc.SetHeaderText(docx.HeaderDefault, "Combined Header"); err != nil {
		t.Fatal(err)
	}
	if err := doc.SetHeaderAlignment(docx.HeaderDefault, "center"); err != nil {
		t.Fatal(err)
	}
	if err := doc.SetFooterText(docx.FooterDefault, "Page "); err != nil {
		t.Fatal(err)
	}
	if err := doc.AddPageNumberAligned(docx.PageNumberArabic, "end"); err != nil {
		t.Fatal(err)
	}

	outFile, err := os.Create(updatedPath)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := doc.WriteTo(outFile); err != nil {
		_ = outFile.Close()
		t.Fatal(err)
	}
	_ = outFile.Close()

	entries, err := readZipEntries(updatedPath)
	if err != nil {
		t.Fatal(err)
	}
	docXML := string(entries["word/document.xml"])
	headerXML := string(entries["word/header_default.xml"])
	footerXML := string(entries["word/footer_default.xml"])
	if !strings.Contains(docXML, "Codex") {
		t.Fatalf("expected placeholder replacement in document.xml, got: %s", docXML)
	}
	if !strings.Contains(docXML, "extNode") {
		t.Fatalf("expected unknown node preserved in document.xml, got: %s", docXML)
	}
	if !strings.Contains(docXML, `w:pgNumType w:fmt="decimal"`) {
		t.Fatalf("expected pgNumType in document.xml, got: %s", docXML)
	}
	if !strings.Contains(headerXML, `w:jc w:val="center"`) {
		t.Fatalf("expected centered header paragraph, got: %s", headerXML)
	}
	if !strings.Contains(footerXML, `w:jc w:val="end"`) {
		t.Fatalf("expected end aligned footer paragraph, got: %s", footerXML)
	}
	if !strings.Contains(footerXML, `fldCharType="begin"`) || !strings.Contains(footerXML, `fldCharType="end"`) {
		t.Fatalf("expected PAGE field boundary in footer xml, got: %s", footerXML)
	}

	if err := roundTripOne(updatedPath, outPath); err != nil {
		t.Fatal(err)
	}
	issues, err := compareDocxStructure(updatedPath, outPath)
	if err != nil {
		t.Fatal(err)
	}
	if len(issues) != 0 {
		t.Fatalf("unexpected structural issues after combined update and round-trip: %v", issues)
	}
}

func createSampleDocxWithUnknownNodes(path string) error {
	base := docx.New().WithDefaultTheme().WithA4Page()
	base.AddParagraph().AddText("roundtrip")

	var buf bytes.Buffer
	if _, err := base.WriteTo(&buf); err != nil {
		return err
	}

	zr, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		return err
	}

	entries := make(map[string][]byte, len(zr.File))
	for _, f := range zr.File {
		rc, err := f.Open()
		if err != nil {
			return err
		}
		data, err := io.ReadAll(rc)
		_ = rc.Close()
		if err != nil {
			return err
		}
		entries[f.Name] = data
	}

	entries["word/document.xml"] = []byte(`<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>before</w:t></w:r></w:p>
    <w:sdt><w:sdtContent><w:p><w:r><w:t>inside</w:t></w:r></w:p></w:sdtContent></w:sdt>
    <w:tbl>
      <w:tblPr/>
      <w:tr>
        <w:tc>
          <w:tcPr/>
          <w:p><w:r><w:t>a</w:t></w:r></w:p>
          <w:extNode foo="bar"/>
          <w:p><w:r><w:t>b</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
  </w:body>
</w:document>`)

	out, err := os.Create(path)
	if err != nil {
		return err
	}
	defer out.Close()

	zw := zip.NewWriter(out)
	for name, data := range entries {
		w, err := zw.Create(name)
		if err != nil {
			return err
		}
		if _, err := w.Write(data); err != nil {
			return err
		}
	}
	return zw.Close()
}

func createSampleDocxWithFieldCodes(path string) error {
	base := docx.New().WithDefaultTheme().WithA4Page()
	base.AddParagraph().AddText("field replacement baseline")

	var buf bytes.Buffer
	if _, err := base.WriteTo(&buf); err != nil {
		return err
	}

	zr, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		return err
	}

	entries := make(map[string][]byte, len(zr.File))
	for _, f := range zr.File {
		rc, err := f.Open()
		if err != nil {
			return err
		}
		data, err := io.ReadAll(rc)
		_ = rc.Close()
		if err != nil {
			return err
		}
		entries[f.Name] = data
	}

	entries["word/document.xml"] = []byte(`<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:fldChar w:fldCharType="begin"/></w:r>
      <w:r><w:instrText xml:space="preserve"> MERGEFIELD he</w:instrText></w:r>
      <w:r><w:instrText xml:space="preserve">llo </w:instrText></w:r>
      <w:r><w:fldChar w:fldCharType="separate"/></w:r>
      <w:r><w:t>result</w:t></w:r>
      <w:r><w:fldChar w:fldCharType="end"/></w:r>
    </w:p>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
  </w:body>
</w:document>`)

	out, err := os.Create(path)
	if err != nil {
		return err
	}
	defer out.Close()

	zw := zip.NewWriter(out)
	for name, data := range entries {
		w, err := zw.Create(name)
		if err != nil {
			return err
		}
		if _, err := w.Write(data); err != nil {
			return err
		}
	}
	return zw.Close()
}

func createSampleDocxWithHeaderFooter(path string) error {
	base := docx.New().WithDefaultTheme().WithA4Page()
	base.AddParagraph().AddText("header footer baseline")
	if err := base.SetHeaderText(docx.HeaderDefault, "Initial Header"); err != nil {
		return err
	}
	if err := base.SetFooterText(docx.FooterDefault, "Initial Footer"); err != nil {
		return err
	}

	out, err := os.Create(path)
	if err != nil {
		return err
	}
	defer out.Close()
	_, err = base.WriteTo(out)
	return err
}

func createSampleDocxWithSections(path string) error {
	base := docx.New().WithDefaultTheme().WithA4Page()
	base.AddParagraph().AddText("seed")

	var buf bytes.Buffer
	if _, err := base.WriteTo(&buf); err != nil {
		return err
	}
	zr, err := zip.NewReader(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		return err
	}
	entries := make(map[string][]byte, len(zr.File))
	for _, f := range zr.File {
		rc, err := f.Open()
		if err != nil {
			return err
		}
		data, err := io.ReadAll(rc)
		_ = rc.Close()
		if err != nil {
			return err
		}
		entries[f.Name] = data
	}

	entries["word/document.xml"] = []byte(`<?xml version="1.0" encoding="UTF-8"?>
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
</w:document>`)

	out, err := os.Create(path)
	if err != nil {
		return err
	}
	defer out.Close()
	zw := zip.NewWriter(out)
	for name, data := range entries {
		w, err := zw.Create(name)
		if err != nil {
			return err
		}
		if _, err := w.Write(data); err != nil {
			return err
		}
	}
	return zw.Close()
}

func createSampleDocxWithSectionsAndSettings(path string) error {
	if err := createSampleDocxWithSections(path); err != nil {
		return err
	}
	zr, err := zip.OpenReader(path)
	if err != nil {
		return err
	}
	defer zr.Close()
	entries := make(map[string][]byte, len(zr.File)+1)
	for _, f := range zr.File {
		rc, err := f.Open()
		if err != nil {
			return err
		}
		data, err := io.ReadAll(rc)
		_ = rc.Close()
		if err != nil {
			return err
		}
		entries[f.Name] = data
	}
	entries["word/settings.xml"] = []byte(`<?xml version="1.0" encoding="UTF-8"?><w:settings xmlns:w="` + docx.XMLNS_W + `"/>`)

	out, err := os.Create(path)
	if err != nil {
		return err
	}
	defer out.Close()
	zw := zip.NewWriter(out)
	for name, data := range entries {
		w, err := zw.Create(name)
		if err != nil {
			return err
		}
		if _, err := w.Write(data); err != nil {
			return err
		}
	}
	return zw.Close()
}
