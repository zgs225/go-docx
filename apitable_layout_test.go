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

func TestTableSetRepeatHeaderAPI(t *testing.T) {
	d := New().WithDefaultTheme()
	tbl := d.AddTable(2, 1, 0, nil)

	tbl.SetRepeatHeader(0, true)
	if tbl.TableRows[0].TableRowProperties == nil || tbl.TableRows[0].TableRowProperties.RepeatHeader == nil {
		t.Fatal("expected repeat header enabled on row 0")
	}
	outEnabled, err := marshalXMLString(tbl)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(outEnabled, "tblHeader") {
		t.Fatalf("expected tblHeader in marshaled table xml, got: %s", outEnabled)
	}

	tbl.SetRepeatHeader(0, false)
	if tbl.TableRows[0].TableRowProperties.RepeatHeader != nil {
		t.Fatal("expected repeat header cleared on row 0")
	}
	outDisabled, err := marshalXMLString(tbl)
	if err != nil {
		t.Fatal(err)
	}
	if strings.Contains(outDisabled, "tblHeader") {
		t.Fatalf("expected tblHeader removed in marshaled table xml, got: %s", outDisabled)
	}

	beforeNoop, err := marshalXMLString(tbl)
	if err != nil {
		t.Fatal(err)
	}
	tbl.SetRepeatHeader(9, true)
	afterNoop, err := marshalXMLString(tbl)
	if err != nil {
		t.Fatal(err)
	}
	if beforeNoop != afterNoop {
		t.Fatal("expected out-of-range SetRepeatHeader to be no-op")
	}
}

func TestTableSetLayoutOptionsAPI(t *testing.T) {
	d := New().WithDefaultTheme()
	tbl := d.AddTable(1, 1, 0, nil)
	tbl.SetLayout(TableLayoutOptions{
		Mode:          TableLayoutModeFixed,
		WidthTwips:    7200,
		Justification: "center",
		DefaultCellPadding: &TablePadding{
			Top: 120, Right: 180, Bottom: 240, Left: 300,
		},
	})
	if tbl.TableProperties == nil {
		t.Fatal("expected table properties")
	}
	if tbl.TableProperties.Layout == nil || tbl.TableProperties.Layout.Type != "fixed" {
		t.Fatalf("unexpected layout: %#v", tbl.TableProperties.Layout)
	}
	if tbl.TableProperties.Width == nil || tbl.TableProperties.Width.Type != "dxa" || tbl.TableProperties.Width.W != 7200 {
		t.Fatalf("unexpected width: %#v", tbl.TableProperties.Width)
	}
	if tbl.TableProperties.Justification == nil || tbl.TableProperties.Justification.Val != "center" {
		t.Fatalf("unexpected table justification: %#v", tbl.TableProperties.Justification)
	}
	if tbl.TableProperties.CellMargins == nil || tbl.TableProperties.CellMargins.Left == nil || tbl.TableProperties.CellMargins.Left.W != 300 {
		t.Fatalf("unexpected default padding: %#v", tbl.TableProperties.CellMargins)
	}

	tbl.SetLayout(TableLayoutOptions{
		Mode:       TableLayoutModeAutofit,
		WidthTwips: 0,
	})
	if tbl.TableProperties.Layout == nil || tbl.TableProperties.Layout.Type != "autofit" {
		t.Fatalf("unexpected layout after autofit: %#v", tbl.TableProperties.Layout)
	}
	if tbl.TableProperties.Width == nil || tbl.TableProperties.Width.Type != "auto" {
		t.Fatalf("unexpected width after auto: %#v", tbl.TableProperties.Width)
	}
	if tbl.TableProperties.Justification == nil || tbl.TableProperties.Justification.Val != "center" {
		t.Fatalf("expected existing justification preserved, got: %#v", tbl.TableProperties.Justification)
	}
	if tbl.TableProperties.CellMargins == nil || tbl.TableProperties.CellMargins.Left == nil || tbl.TableProperties.CellMargins.Left.W != 300 {
		t.Fatalf("expected existing padding preserved, got: %#v", tbl.TableProperties.CellMargins)
	}
}

func TestTableSetRowLayoutOptionsAPI(t *testing.T) {
	d := New().WithDefaultTheme()
	tbl := d.AddTable(2, 1, 0, nil)

	repeatTrue := true
	tbl.SetRowLayout(0, RowLayoutOptions{
		Justification: "end",
		HeightTwips:   900,
		HeightRule:    "exact",
		RepeatHeader:  &repeatTrue,
	})
	r0 := tbl.TableRows[0]
	if r0.TableRowProperties == nil || r0.TableRowProperties.Justification == nil || r0.TableRowProperties.Justification.Val != "end" {
		t.Fatalf("unexpected row 0 justification: %#v", r0.TableRowProperties)
	}
	if r0.TableRowProperties.TableRowHeight == nil || r0.TableRowProperties.TableRowHeight.Val != 900 || r0.TableRowProperties.TableRowHeight.Rule != "exact" {
		t.Fatalf("unexpected row 0 height: %#v", r0.TableRowProperties.TableRowHeight)
	}
	if r0.TableRowProperties.RepeatHeader == nil {
		t.Fatal("expected row 0 repeat header enabled")
	}

	tbl.SetRowLayout(0, RowLayoutOptions{
		HeightTwips:  900,
		HeightRule:   "bad-rule",
		RepeatHeader: nil,
	})
	if r0.TableRowProperties.RepeatHeader == nil {
		t.Fatal("expected repeat header unchanged when RepeatHeader=nil")
	}
	if r0.TableRowProperties.TableRowHeight == nil || r0.TableRowProperties.TableRowHeight.Rule != "exact" {
		t.Fatalf("expected height rule unchanged on invalid input, got: %#v", r0.TableRowProperties.TableRowHeight)
	}

	repeatFalse := false
	tbl.SetRowLayout(0, RowLayoutOptions{
		HeightTwips:  900,
		HeightRule:   "exact",
		RepeatHeader: &repeatFalse,
	})
	if r0.TableRowProperties.RepeatHeader != nil {
		t.Fatal("expected repeat header cleared")
	}

	tbl.SetRowLayout(0, RowLayoutOptions{HeightTwips: 0})
	if r0.TableRowProperties.TableRowHeight != nil {
		t.Fatal("expected trHeight cleared when HeightTwips <= 0")
	}

	repeatTrue = true
	tbl.TableRows[1].SetLayout(RowLayoutOptions{
		Justification: "center",
		HeightTwips:   1200,
		HeightRule:    "atLeast",
		RepeatHeader:  &repeatTrue,
	})
	r1 := tbl.TableRows[1]
	if r1.TableRowProperties == nil || r1.TableRowProperties.TableRowHeight == nil || r1.TableRowProperties.TableRowHeight.Rule != "atLeast" {
		t.Fatalf("unexpected row 1 layout: %#v", r1.TableRowProperties)
	}

	beforeNoop, err := marshalXMLString(tbl)
	if err != nil {
		t.Fatal(err)
	}
	tbl.SetRowLayout(99, RowLayoutOptions{HeightTwips: 500, HeightRule: "exact"})
	afterNoop, err := marshalXMLString(tbl)
	if err != nil {
		t.Fatal(err)
	}
	if beforeNoop != afterNoop {
		t.Fatal("expected out-of-range SetRowLayout to be no-op")
	}
}

func TestTableSetRowLayoutClearSwitchKeepsExistingJustification(t *testing.T) {
	d := New().WithDefaultTheme()
	tbl := d.AddTable(1, 1, 0, nil)

	repeatTrue := true
	tbl.SetRowLayout(0, RowLayoutOptions{
		Justification: "center",
		HeightTwips:   900,
		HeightRule:    "exact",
		RepeatHeader:  &repeatTrue,
	})
	row := tbl.TableRows[0]
	if row.TableRowProperties == nil || row.TableRowProperties.Justification == nil || row.TableRowProperties.Justification.Val != "center" {
		t.Fatalf("expected initial row justification center, got: %#v", row.TableRowProperties)
	}
	if row.TableRowProperties.RepeatHeader == nil || row.TableRowProperties.TableRowHeight == nil {
		t.Fatalf("expected initial repeat header and height, got: %#v", row.TableRowProperties)
	}

	repeatFalse := false
	tbl.SetRowLayout(0, RowLayoutOptions{
		HeightTwips:  0,
		RepeatHeader: &repeatFalse,
	})
	if row.TableRowProperties.RepeatHeader != nil {
		t.Fatal("expected repeat header cleared")
	}
	if row.TableRowProperties.TableRowHeight != nil {
		t.Fatal("expected trHeight cleared")
	}
	if row.TableRowProperties.Justification == nil || row.TableRowProperties.Justification.Val != "center" {
		t.Fatalf("expected existing row justification preserved, got: %#v", row.TableRowProperties.Justification)
	}
}

func TestTableRepeatHeaderAndLayoutRoundTripStable(t *testing.T) {
	d := New().WithDefaultTheme()
	tbl := d.AddTable(2, 2, 0, nil)
	tbl.SetLayout(TableLayoutOptions{
		Mode:       TableLayoutModeFixed,
		WidthTwips: 7200,
		DefaultCellPadding: &TablePadding{
			Top: 120, Right: 180, Bottom: 240, Left: 300,
		},
	})
	tbl.SetRepeatHeader(0, true)
	tbl.SetRowLayout(1, RowLayoutOptions{
		Justification: "center",
		HeightTwips:   800,
		HeightRule:    "atLeast",
	})
	tbl.TableRows[0].TableCells[0].
		SetColSpan(2).
		SetCellBordersSame("single", 8, 0, "000000").
		Padding(10, 20, 30, 40)

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
	for _, needle := range []string{"tblHeader", "trHeight", "tblLayout", "tblCellMar", "gridSpan", "tcBorders", "tcMar"} {
		if !strings.Contains(out1, needle) || !strings.Contains(out2, needle) {
			t.Fatalf("expected %q in both marshaled outputs", needle)
		}
	}
}
