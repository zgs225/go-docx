package docx

import (
	"bytes"
	"encoding/xml"
	"os"
	"strings"
	"testing"
)

func TestTableCellSetColSpan(t *testing.T) {
	d := New().WithDefaultTheme()
	cell := d.AddTable(1, 1, 0, nil).TableRows[0].TableCells[0]

	cell.SetColSpan(3)
	if cell.TableCellProperties == nil || cell.TableCellProperties.GridSpan == nil || cell.TableCellProperties.GridSpan.Val != 3 {
		t.Fatalf("expected gridSpan=3, got %#v", cell.TableCellProperties)
	}

	cell.SetColSpan(1)
	if cell.TableCellProperties.GridSpan != nil {
		t.Fatal("expected gridSpan cleared when cols <= 1")
	}
}

func TestTableCellSetColSpanBoundary(t *testing.T) {
	d := New().WithDefaultTheme()
	cell := d.AddTable(1, 1, 0, nil).TableRows[0].TableCells[0]

	cell.SetColSpan(2)
	if cell.TableCellProperties == nil || cell.TableCellProperties.GridSpan == nil || cell.TableCellProperties.GridSpan.Val != 2 {
		t.Fatalf("expected gridSpan=2, got %#v", cell.TableCellProperties)
	}

	cell.SetColSpan(0)
	if cell.TableCellProperties.GridSpan != nil {
		t.Fatalf("expected gridSpan cleared by 0, got %#v", cell.TableCellProperties.GridSpan)
	}

	cell.SetColSpan(-1)
	if cell.TableCellProperties.GridSpan != nil {
		t.Fatalf("expected gridSpan cleared by negative value, got %#v", cell.TableCellProperties.GridSpan)
	}
}

func TestTableCellSetRowSpan(t *testing.T) {
	d := New().WithDefaultTheme()
	cell := d.AddTable(1, 1, 0, nil).TableRows[0].TableCells[0]

	cell.SetRowSpanRestart()
	if cell.TableCellProperties == nil || cell.TableCellProperties.VMerge == nil || cell.TableCellProperties.VMerge.Val != "restart" {
		t.Fatalf("expected vMerge restart, got %#v", cell.TableCellProperties)
	}

	cell.SetRowSpanContinue()
	if cell.TableCellProperties.VMerge == nil || cell.TableCellProperties.VMerge.Val != "" {
		t.Fatalf("expected vMerge continue (empty val), got %#v", cell.TableCellProperties.VMerge)
	}

	cell.ClearRowSpan()
	if cell.TableCellProperties.VMerge != nil {
		t.Fatal("expected vMerge cleared")
	}
}

func TestTableCellSetBordersAndPaddingCoexist(t *testing.T) {
	d := New().WithDefaultTheme()
	cell := d.AddTable(1, 1, 0, nil).TableRows[0].TableCells[0]

	cell.Padding(120, 180, 240, 300).
		SetCellBordersSame("single", 8, 0, "FF0000")

	if cell.TableCellProperties == nil || cell.TableCellProperties.Margins == nil {
		t.Fatal("expected tcMar to exist")
	}
	if cell.TableCellProperties.TableBorders == nil {
		t.Fatal("expected tcBorders to exist")
	}
	if cell.TableCellProperties.TableBorders.Top == nil || cell.TableCellProperties.TableBorders.Top.Color != "FF0000" {
		t.Fatalf("unexpected top border: %#v", cell.TableCellProperties.TableBorders.Top)
	}
	if cell.TableCellProperties.Margins.Left == nil || cell.TableCellProperties.Margins.Left.W != 300 {
		t.Fatalf("unexpected left padding: %#v", cell.TableCellProperties.Margins.Left)
	}

	cell.SetCellBorderLeft("double", 12, 1, "00FF00")
	if cell.TableCellProperties.TableBorders.Left == nil || cell.TableCellProperties.TableBorders.Left.Val != "double" {
		t.Fatalf("unexpected left border after override: %#v", cell.TableCellProperties.TableBorders.Left)
	}

	cell.ClearCellBorders()
	if cell.TableCellProperties.TableBorders != nil {
		t.Fatal("expected tcBorders cleared")
	}
}

func TestTableCellSetCellBorderEachSide(t *testing.T) {
	d := New().WithDefaultTheme()
	cell := d.AddTable(1, 1, 0, nil).TableRows[0].TableCells[0]

	cell.SetCellBorderTop("single", 8, 0, "111111")
	cell.SetCellBorderRight("double", 12, 1, "222222")
	cell.SetCellBorderBottom("dashed", 16, 2, "333333")
	cell.SetCellBorderLeft("dotted", 20, 3, "444444")

	b := cell.TableCellProperties.TableBorders
	if b == nil {
		t.Fatal("expected tcBorders to exist")
	}
	if b.Top == nil || b.Top.Val != "single" || b.Top.Size != 8 || b.Top.Space != 0 || b.Top.Color != "111111" {
		t.Fatalf("unexpected top border: %#v", b.Top)
	}
	if b.Right == nil || b.Right.Val != "double" || b.Right.Size != 12 || b.Right.Space != 1 || b.Right.Color != "222222" {
		t.Fatalf("unexpected right border: %#v", b.Right)
	}
	if b.Bottom == nil || b.Bottom.Val != "dashed" || b.Bottom.Size != 16 || b.Bottom.Space != 2 || b.Bottom.Color != "333333" {
		t.Fatalf("unexpected bottom border: %#v", b.Bottom)
	}
	if b.Left == nil || b.Left.Val != "dotted" || b.Left.Size != 20 || b.Left.Space != 3 || b.Left.Color != "444444" {
		t.Fatalf("unexpected left border: %#v", b.Left)
	}
}

func TestTableCellSetCellBordersSameThenOverride(t *testing.T) {
	d := New().WithDefaultTheme()
	cell := d.AddTable(1, 1, 0, nil).TableRows[0].TableCells[0]

	cell.SetCellBordersSame("single", 8, 0, "AAAAAA")
	cell.SetCellBorderRight("double", 12, 1, "BBBBBB")

	b := cell.TableCellProperties.TableBorders
	if b == nil {
		t.Fatal("expected tcBorders to exist")
	}
	if b.Top == nil || b.Top.Val != "single" || b.Top.Color != "AAAAAA" {
		t.Fatalf("unexpected top border: %#v", b.Top)
	}
	if b.Right == nil || b.Right.Val != "double" || b.Right.Size != 12 || b.Right.Space != 1 || b.Right.Color != "BBBBBB" {
		t.Fatalf("unexpected overridden right border: %#v", b.Right)
	}
	if b.Bottom == nil || b.Bottom.Val != "single" || b.Bottom.Color != "AAAAAA" {
		t.Fatalf("unexpected bottom border: %#v", b.Bottom)
	}
	if b.Left == nil || b.Left.Val != "single" || b.Left.Color != "AAAAAA" {
		t.Fatalf("unexpected left border: %#v", b.Left)
	}
}

func TestTableCellClearRowSpanKeepsOtherTcPrFields(t *testing.T) {
	d := New().WithDefaultTheme()
	cell := d.AddTable(1, 1, 0, nil).TableRows[0].TableCells[0]

	cell.SetColSpan(3).
		Padding(120, 180, 240, 300).
		SetCellBordersSame("single", 8, 0, "000000").
		SetRowSpanRestart().
		ClearRowSpan()

	tcpr := cell.TableCellProperties
	if tcpr == nil {
		t.Fatal("expected tcPr to exist")
	}
	if tcpr.VMerge != nil {
		t.Fatalf("expected vMerge cleared, got %#v", tcpr.VMerge)
	}
	if tcpr.GridSpan == nil || tcpr.GridSpan.Val != 3 {
		t.Fatalf("expected gridSpan kept, got %#v", tcpr.GridSpan)
	}
	if tcpr.Margins == nil || tcpr.Margins.Left == nil || tcpr.Margins.Left.W != 300 {
		t.Fatalf("expected margins kept, got %#v", tcpr.Margins)
	}
	if tcpr.TableBorders == nil || tcpr.TableBorders.Top == nil || tcpr.TableBorders.Top.Val != "single" {
		t.Fatalf("expected borders kept, got %#v", tcpr.TableBorders)
	}
}

func TestTableCellClearCellBordersKeepsMarginsAndSpan(t *testing.T) {
	d := New().WithDefaultTheme()
	cell := d.AddTable(1, 1, 0, nil).TableRows[0].TableCells[0]

	cell.SetColSpan(2).
		SetRowSpanContinue().
		Padding(120, 180, 240, 300).
		SetCellBordersSame("single", 8, 0, "000000").
		ClearCellBorders()

	tcpr := cell.TableCellProperties
	if tcpr == nil {
		t.Fatal("expected tcPr to exist")
	}
	if tcpr.TableBorders != nil {
		t.Fatalf("expected tcBorders cleared, got %#v", tcpr.TableBorders)
	}
	if tcpr.GridSpan == nil || tcpr.GridSpan.Val != 2 {
		t.Fatalf("expected gridSpan kept, got %#v", tcpr.GridSpan)
	}
	if tcpr.VMerge == nil || tcpr.VMerge.Val != "" {
		t.Fatalf("expected vMerge continue kept, got %#v", tcpr.VMerge)
	}
	if tcpr.Margins == nil || tcpr.Margins.Top == nil || tcpr.Margins.Top.W != 120 {
		t.Fatalf("expected margins kept, got %#v", tcpr.Margins)
	}
}

func TestTableCellRowSpanContinueSerializesWithoutVal(t *testing.T) {
	d := New().WithDefaultTheme()
	cell := d.AddTable(1, 1, 0, nil).TableRows[0].TableCells[0]
	cell.SetRowSpanContinue()

	out, err := marshalXMLString(cell)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(out, "vMerge") {
		t.Fatalf("expected vMerge element, got: %s", out)
	}
	if strings.Contains(out, `vMerge w:val=`) || strings.Contains(out, `w:vMerge w:val=`) {
		t.Fatalf("continue form should not contain w:val, got: %s", out)
	}
}

func TestTableCellEnsureTcPrInsertedInOrderedPath(t *testing.T) {
	const in = `<w:tc xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>a</w:t></w:r></w:p><w:extNode foo="bar"/></w:tc>`
	var cell WTableCell
	if err := xml.Unmarshal(StringToBytes(in), &cell); err != nil {
		t.Fatal(err)
	}
	cell.SetColSpan(2).
		SetRowSpanRestart().
		SetCellBorderTop("single", 8, 0, "000000")

	out, err := marshalXMLString(&cell)
	if err != nil {
		t.Fatal(err)
	}

	idxTcPr := strings.Index(out, "<w:tcPr")
	idxP := strings.Index(out, "<w:p")
	if idxTcPr < 0 || idxP < 0 || idxTcPr > idxP {
		t.Fatalf("expected tcPr before first paragraph, got: %s", out)
	}
	if !strings.Contains(out, "extNode") {
		t.Fatalf("expected unknown node preserved, got: %s", out)
	}
	if !strings.Contains(out, "gridSpan") || !strings.Contains(out, "vMerge") || !strings.Contains(out, "tcBorders") {
		t.Fatalf("expected tcPr to include span/merge/border nodes, got: %s", out)
	}
	idxExt := strings.Index(out, "extNode")
	if idxExt < 0 || idxExt < idxP {
		t.Fatalf("expected unknown node after paragraph order preserved, got: %s", out)
	}
}

func TestTableSpanBorderRoundTripAndUnknownNodes(t *testing.T) {
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

	var target *WTableCell
	for _, item := range d.Document.Body.Items {
		tbl, ok := item.(*Table)
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
		Padding(120, 120, 120, 120).
		SetCellBordersSame("single", 8, 0, "000000")

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
		t.Fatal("expected preserved RawXMLNode in body after table span/border update and round-trip")
	}
}

func TestTableSpanBorderSetThenClearRoundTrip(t *testing.T) {
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

	var target *WTableCell
	for _, item := range d.Document.Body.Items {
		tbl, ok := item.(*Table)
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
		Padding(120, 120, 120, 120).
		SetCellBordersSame("single", 8, 0, "000000").
		ClearCellBorders()

	var buf bytes.Buffer
	if _, err := d.WriteTo(&buf); err != nil {
		t.Fatal(err)
	}
	d2, err := Parse(bytes.NewReader(buf.Bytes()), int64(buf.Len()))
	if err != nil {
		t.Fatal(err)
	}

	var got *WTableCell
	for _, item := range d2.Document.Body.Items {
		tbl, ok := item.(*Table)
		if !ok || len(tbl.TableRows) == 0 || len(tbl.TableRows[0].TableCells) == 0 {
			continue
		}
		got = tbl.TableRows[0].TableCells[0]
		break
	}
	if got == nil || got.TableCellProperties == nil {
		t.Fatalf("expected parsed table cell properties, got %#v", got)
	}
	if got.TableCellProperties.TableBorders != nil {
		t.Fatalf("expected cleared borders after round-trip, got %#v", got.TableCellProperties.TableBorders)
	}
	if got.TableCellProperties.GridSpan == nil || got.TableCellProperties.GridSpan.Val != 2 {
		t.Fatalf("expected gridSpan kept after round-trip, got %#v", got.TableCellProperties.GridSpan)
	}
	if got.TableCellProperties.VMerge == nil || got.TableCellProperties.VMerge.Val != "" {
		t.Fatalf("expected continue vMerge kept after round-trip, got %#v", got.TableCellProperties.VMerge)
	}
	if got.TableCellProperties.Margins == nil || got.TableCellProperties.Margins.Top == nil || got.TableCellProperties.Margins.Top.W != 120 {
		t.Fatalf("expected margins kept after round-trip, got %#v", got.TableCellProperties.Margins)
	}
}
