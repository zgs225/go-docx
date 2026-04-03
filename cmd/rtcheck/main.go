package main

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"encoding/xml"
	"errors"
	"flag"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"sort"
	"strings"

	"github.com/zgs225/go-docx"
)

type fileResult struct {
	File   string   `json:"file"`
	Pass   bool     `json:"pass"`
	Issues []string `json:"issues,omitempty"`
}

type report struct {
	InputDir  string       `json:"input_dir"`
	OutputDir string       `json:"output_dir"`
	Passed    int          `json:"passed"`
	Failed    int          `json:"failed"`
	Results   []fileResult `json:"results"`
}

func main() {
	inDir := flag.String("in", "", "input directory containing .docx samples")
	outDir := flag.String("out", "", "output directory to store round-tripped .docx")
	reportPath := flag.String("report", "", "json report path")
	flag.Parse()

	if *inDir == "" || *outDir == "" || *reportPath == "" {
		fmt.Fprintln(os.Stderr, "usage: go run ./cmd/rtcheck --in <samples_dir> --out <tmp_dir> --report <report.json>")
		os.Exit(2)
	}

	if err := os.MkdirAll(*outDir, 0o755); err != nil {
		fmt.Fprintln(os.Stderr, "prepare out dir:", err)
		os.Exit(2)
	}

	entries, err := os.ReadDir(*inDir)
	if err != nil {
		fmt.Fprintln(os.Stderr, "read input dir:", err)
		os.Exit(2)
	}

	rep := report{
		InputDir:  *inDir,
		OutputDir: *outDir,
		Results:   make([]fileResult, 0, len(entries)),
	}

	for _, ent := range entries {
		if ent.IsDir() || !strings.HasSuffix(strings.ToLower(ent.Name()), ".docx") {
			continue
		}
		inPath := filepath.Join(*inDir, ent.Name())
		outPath := filepath.Join(*outDir, ent.Name())

		res := fileResult{File: ent.Name(), Pass: true}
		if err := roundTripOne(inPath, outPath); err != nil {
			res.Pass = false
			res.Issues = append(res.Issues, "round-trip failed: "+err.Error())
		} else {
			issues, err := compareDocxStructure(inPath, outPath)
			if err != nil {
				res.Pass = false
				res.Issues = append(res.Issues, "compare failed: "+err.Error())
			} else if len(issues) > 0 {
				res.Pass = false
				res.Issues = append(res.Issues, issues...)
			}
		}

		if res.Pass {
			rep.Passed++
			fmt.Printf("[PASS] %s\n", ent.Name())
		} else {
			rep.Failed++
			fmt.Printf("[FAIL] %s\n", ent.Name())
			for _, issue := range res.Issues {
				fmt.Printf("  - %s\n", issue)
			}
		}
		rep.Results = append(rep.Results, res)
	}

	data, err := json.MarshalIndent(rep, "", "  ")
	if err != nil {
		fmt.Fprintln(os.Stderr, "marshal report:", err)
		os.Exit(2)
	}
	if err := os.WriteFile(*reportPath, data, 0o644); err != nil {
		fmt.Fprintln(os.Stderr, "write report:", err)
		os.Exit(2)
	}

	if rep.Failed > 0 {
		os.Exit(1)
	}
}

func roundTripOne(inPath, outPath string) error {
	f, err := os.Open(inPath)
	if err != nil {
		return err
	}
	defer f.Close()

	st, err := f.Stat()
	if err != nil {
		return err
	}
	doc, err := docx.Parse(f, st.Size())
	if err != nil {
		return err
	}

	out, err := os.Create(outPath)
	if err != nil {
		return err
	}
	defer out.Close()
	_, err = doc.WriteTo(out)
	return err
}

func compareDocxStructure(aPath, bPath string) ([]string, error) {
	a, err := readZipEntries(aPath)
	if err != nil {
		return nil, err
	}
	b, err := readZipEntries(bPath)
	if err != nil {
		return nil, err
	}

	files := make([]string, 0, len(a))
	for name := range a {
		if isKeyXML(name) {
			files = append(files, name)
		}
	}
	sort.Strings(files)

	issues := make([]string, 0, 8)
	for _, name := range files {
		bData, ok := b[name]
		if !ok {
			issues = append(issues, fmt.Sprintf("missing key file in output: %s", name))
			continue
		}
		same, reason, err := xmlStructurallyEqual(a[name], bData)
		if err != nil {
			issues = append(issues, fmt.Sprintf("%s compare error: %v", name, err))
			continue
		}
		if !same {
			issues = append(issues, fmt.Sprintf("%s mismatch: %s", name, reason))
		}
	}
	if len(files) == 0 {
		return nil, errors.New("no key xml files found in input")
	}
	return issues, nil
}

func readZipEntries(path string) (map[string][]byte, error) {
	zr, err := zip.OpenReader(path)
	if err != nil {
		return nil, err
	}
	defer zr.Close()

	m := make(map[string][]byte, len(zr.File))
	for _, f := range zr.File {
		rc, err := f.Open()
		if err != nil {
			return nil, err
		}
		data, err := io.ReadAll(rc)
		_ = rc.Close()
		if err != nil {
			return nil, err
		}
		m[f.Name] = data
	}
	return m, nil
}

func isKeyXML(name string) bool {
	if !strings.HasSuffix(name, ".xml") && !strings.HasSuffix(name, ".rels") {
		return false
	}
	return strings.HasPrefix(name, "word/") || strings.HasPrefix(name, "_rels/") || strings.Contains(name, "/_rels/")
}

func xmlStructurallyEqual(a, b []byte) (bool, string, error) {
	aSig, err := canonicalXMLSignature(a)
	if err != nil {
		return false, "", err
	}
	bSig, err := canonicalXMLSignature(b)
	if err != nil {
		return false, "", err
	}
	if len(aSig) != len(bSig) {
		return false, fmt.Sprintf("token length %d != %d", len(aSig), len(bSig)), nil
	}
	for i := range aSig {
		if aSig[i] != bSig[i] {
			return false, fmt.Sprintf("first diff at token %d: %q != %q", i, aSig[i], bSig[i]), nil
		}
	}
	return true, "", nil
}

func canonicalXMLSignature(data []byte) ([]string, error) {
	dec := xml.NewDecoder(bytes.NewReader(data))
	sig := make([]string, 0, 256)
	for {
		tok, err := dec.Token()
		if err == io.EOF {
			return sig, nil
		}
		if err != nil {
			return nil, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			attrs := make([]string, 0, len(t.Attr))
			for _, a := range t.Attr {
				if isNSDeclAttr(a) {
					continue
				}
				attrs = append(attrs, nameKey(a.Name)+"="+a.Value)
			}
			sort.Strings(attrs)
			sig = append(sig, "S:"+nameKey(t.Name)+"|"+strings.Join(attrs, ","))
		case xml.EndElement:
			sig = append(sig, "E:"+nameKey(t.Name))
		case xml.CharData:
			text := strings.TrimSpace(string(t))
			if text != "" {
				sig = append(sig, "T:"+text)
			}
		}
	}
}

func isNSDeclAttr(a xml.Attr) bool {
	if a.Name.Space == "xmlns" {
		return true
	}
	return a.Name.Space == "" && strings.HasPrefix(a.Name.Local, "xmlns")
}

func nameKey(n xml.Name) string {
	if n.Space == "" {
		return n.Local
	}
	return n.Space + ":" + n.Local
}
