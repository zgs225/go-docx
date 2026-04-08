package main

import (
	"archive/zip"
	"bytes"
	"io"
	"os"
	"path/filepath"
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
