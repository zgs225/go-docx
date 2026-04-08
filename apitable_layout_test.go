package docx

import (
	"bytes"
	"encoding/xml"
	"os"
	"strings"
	"testing"
)

func TestTablePropertiesUnmarshalCellMarginsAndLayout(t *testing.T) {
	const in = `<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:tblPr>
  <w:tblW w:w="7200" w:type="dxa"/>
  <w:tblLayout w:type="fixed"/>
  <w:tblCellMar>
    <w:top w:w="120" w:type="dxa"/>
    <w:right w:w="180" w:type="dxa"/>
    <w:bottom w:w="240" w:type="dxa"/>
    <w:left w:w="300" w:type="dxa"/>
  </w:tblCellMar>
</w:tblPr>
<w:tr><w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc></w:tr>
</w:tbl>`
	var tbl Table
	if err := xml.Unmarshal(StringToBytes(in), &tbl); err != nil {
		t.Fatal(err)
	}
	if tbl.TableProperties == nil {
		t.Fatal("expected table properties")
	}
	if tbl.TableProperties.Layout == nil || tbl.TableProperties.Layout.Type != "fixed" {
		t.Fatalf("unexpected layout: %#v", tbl.TableProperties.Layout)
	}
	if tbl.TableProperties.CellMargins == nil || tbl.TableProperties.CellMargins.Left == nil || tbl.TableProperties.CellMargins.Left.W != 300 {
		t.Fatalf("unexpected table cell margins: %#v", tbl.TableProperties.CellMargins)
	}
}

func TestTableSetDefaultCellPaddingAndLayoutAPIs(t *testing.T) {
	d := New().WithDefaultTheme()
	tbl := d.AddTable(1, 1, 0, nil)
	tbl.SetDefaultCellPadding(120, 180, 240, 300).SetLayoutFixed().SetWidthTwips(7200)

	if tbl.TableProperties == nil {
		t.Fatal("expected table properties")
	}
	if tbl.TableProperties.CellMargins == nil || tbl.TableProperties.CellMargins.Top == nil || tbl.TableProperties.CellMargins.Top.W != 120 {
		t.Fatalf("unexpected default padding: %#v", tbl.TableProperties.CellMargins)
	}
	if tbl.TableProperties.Layout == nil || tbl.TableProperties.Layout.Type != "fixed" {
		t.Fatalf("unexpected layout: %#v", tbl.TableProperties.Layout)
	}
	if tbl.TableProperties.Width == nil || tbl.TableProperties.Width.W != 7200 || tbl.TableProperties.Width.Type != "dxa" {
		t.Fatalf("unexpected width: %#v", tbl.TableProperties.Width)
	}

	tbl.SetLayoutAutofit().SetWidthTwips(0)
	if tbl.TableProperties.Layout == nil || tbl.TableProperties.Layout.Type != "autofit" {
		t.Fatalf("unexpected layout after autofit: %#v", tbl.TableProperties.Layout)
	}
	if tbl.TableProperties.Width == nil || tbl.TableProperties.Width.Type != "auto" {
		t.Fatalf("unexpected width after reset: %#v", tbl.TableProperties.Width)
	}
}

func TestTableDefaultPaddingAndCellPaddingCoexist(t *testing.T) {
	d := New().WithDefaultTheme()
	tbl := d.AddTable(1, 1, 0, nil)
	tbl.SetDefaultCellPadding(120, 180, 240, 300)
	cell := tbl.TableRows[0].TableCells[0]
	cell.Padding(10, 20, 30, 40)

	if tbl.TableProperties == nil || tbl.TableProperties.CellMargins == nil || tbl.TableProperties.CellMargins.Left == nil || tbl.TableProperties.CellMargins.Left.W != 300 {
		t.Fatalf("unexpected table default padding: %#v", tbl.TableProperties)
	}
	if cell.TableCellProperties == nil || cell.TableCellProperties.Margins == nil || cell.TableCellProperties.Margins.Left == nil || cell.TableCellProperties.Margins.Left.W != 40 {
		t.Fatalf("unexpected cell padding: %#v", cell.TableCellProperties)
	}
}

func TestTableLayoutAndDefaultPaddingRoundTripStable(t *testing.T) {
	d := New().WithDefaultTheme()
	tbl := d.AddTableTwips([]int64{1200, 1200}, []int64{2400, 2400}, 0, nil)
	tbl.SetDefaultCellPadding(120, 180, 240, 300).SetLayoutFixed().SetWidthTwips(7200)
	tbl.TableRows[0].TableCells[0].Padding(10, 20, 30, 40)

	out1, err := marshalXMLString(tbl)
	if err != nil {
		t.Fatal(err)
	}

	var tbl2 Table
	if err := xml.Unmarshal(StringToBytes(out1), &tbl2); err != nil {
		t.Fatal(err)
	}
	out2, err := marshalXMLString(&tbl2)
	if err != nil {
		t.Fatal(err)
	}

	for _, needle := range []string{"tblLayout", "tblCellMar", "tcMar"} {
		if !strings.Contains(out1, needle) || !strings.Contains(out2, needle) {
			t.Fatalf("expected %q in both marshaled outputs", needle)
		}
	}
}

func TestTableLayoutDefaultPaddingWithUnknownNodesRoundTrip(t *testing.T) {
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

	var target *Table
	for _, item := range d.Document.Body.Items {
		if tbl, ok := item.(*Table); ok {
			target = tbl
			break
		}
	}
	if target == nil {
		t.Fatal("expected at least one table in sample")
	}

	target.SetDefaultCellPadding(120, 180, 240, 300).SetLayoutFixed().SetWidthTwips(7200)
	target.TableRows[0].TableCells[0].Padding(10, 20, 30, 40)

	var buf bytes.Buffer
	if _, err := d.WriteTo(&buf); err != nil {
		t.Fatal(err)
	}
	d2, err := Parse(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
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
		t.Fatal("expected preserved RawXMLNode in body after table layout/padding update and round-trip")
	}
}
