package main

import (
	"archive/zip"
	"bytes"
	"io"
	"os"
	"path/filepath"
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
