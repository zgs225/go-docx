/*
   Copyright (c) 2020 gingfrederik
   Copyright (c) 2021 Gonzalo Fernandez-Victorio
   Copyright (c) 2021 Basement Crowd Ltd (https://www.basementcrowd.com)
   Copyright (c) 2023 Fumiama Minamoto (源文雨)

   This program is free software: you can redistribute it and/or modify
   it under the terms of the GNU Affero General Public License as published
   by the Free Software Foundation, either version 3 of the License, or
   (at your option) any later version.

   This program is distributed in the hope that it will be useful,
   but WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
   GNU Affero General Public License for more details.

   You should have received a copy of the GNU Affero General Public License
   along with this program.  If not, see <https://www.gnu.org/licenses/>.
*/

package docx

import (
	"reflect"
)

// AddTable add a new table to body by col*row
//
// unit: twips (1/20 point)
func (f *Docx) AddTable(
	row int,
	col int,
	tableWidth int64,
	borderColors *APITableBorderColors,
) *Table {
	trs := make([]*WTableRow, row)
	for i := 0; i < row; i++ {
		cells := make([]*WTableCell, col)
		for i := range cells {
			cells[i] = &WTableCell{
				TableCellProperties: &WTableCellProperties{
					TableCellWidth: &WTableCellWidth{Type: "auto"},
				},
				file: f,
			}
		}
		trs[i] = &WTableRow{
			TableRowProperties: &WTableRowProperties{},
			TableCells:         cells,
		}
	}

	if borderColors == nil {
		borderColors = new(APITableBorderColors)
	}
	borderColors.applyDefault()

	wTableWidth := &WTableWidth{Type: "auto"}

	if tableWidth > 0 {
		wTableWidth = &WTableWidth{W: tableWidth}
	}

	tbl := &Table{
		TableProperties: &WTableProperties{
			Width: wTableWidth,
			TableBorders: &WTableBorders{
				Top:     &WTableBorder{Val: "single", Size: 4, Space: 0, Color: borderColors.Top},
				Left:    &WTableBorder{Val: "single", Size: 4, Space: 0, Color: borderColors.Left},
				Bottom:  &WTableBorder{Val: "single", Size: 4, Space: 0, Color: borderColors.Bottom},
				Right:   &WTableBorder{Val: "single", Size: 4, Space: 0, Color: borderColors.Right},
				InsideH: &WTableBorder{Val: "single", Size: 4, Space: 0, Color: borderColors.InsideH},
				InsideV: &WTableBorder{Val: "single", Size: 4, Space: 0, Color: borderColors.InsideV},
			},
			Look: &WTableLook{
				Val: "0000",
			},
		},
		TableGrid: &WTableGrid{},
		TableRows: trs,
	}
	f.Document.Body.Items = append(f.Document.Body.Items, tbl)
	return tbl
}

// AddTableTwips add a new table to body by height and width
//
// unit: twips (1/20 point)
func (f *Docx) AddTableTwips(
	rowHeights []int64,
	colWidths []int64,
	tableWidth int64,
	borderColors *APITableBorderColors,
) *Table {
	grids := make([]*WGridCol, len(colWidths))
	trs := make([]*WTableRow, len(rowHeights))
	for i, w := range colWidths {
		if w > 0 {
			grids[i] = &WGridCol{
				W: w,
			}
		}
	}
	for i, h := range rowHeights {
		cells := make([]*WTableCell, len(colWidths))
		for i, w := range colWidths {
			cells[i] = &WTableCell{
				TableCellProperties: &WTableCellProperties{
					TableCellWidth: &WTableCellWidth{W: w, Type: "dxa"},
				},
				file: f,
			}
		}
		trs[i] = &WTableRow{
			TableRowProperties: &WTableRowProperties{},
			TableCells:         cells,
		}
		if h > 0 {
			trs[i].TableRowProperties.TableRowHeight = &WTableRowHeight{
				Val: h,
			}
		}
	}

	if borderColors == nil {
		borderColors = new(APITableBorderColors)
	}
	borderColors.applyDefault()

	wTableWidth := &WTableWidth{Type: "auto"}

	if tableWidth > 0 {
		wTableWidth = &WTableWidth{W: tableWidth}
	}

	tbl := &Table{
		TableProperties: &WTableProperties{
			Width: wTableWidth,
			TableBorders: &WTableBorders{
				Top:     &WTableBorder{Val: "single", Size: 4, Space: 0, Color: borderColors.Top},
				Left:    &WTableBorder{Val: "single", Size: 4, Space: 0, Color: borderColors.Left},
				Bottom:  &WTableBorder{Val: "single", Size: 4, Space: 0, Color: borderColors.Bottom},
				Right:   &WTableBorder{Val: "single", Size: 4, Space: 0, Color: borderColors.Right},
				InsideH: &WTableBorder{Val: "single", Size: 4, Space: 0, Color: borderColors.InsideH},
				InsideV: &WTableBorder{Val: "single", Size: 4, Space: 0, Color: borderColors.InsideV},
			},
			Look: &WTableLook{
				Val: "0000",
			},
		},
		TableGrid: &WTableGrid{
			GridCols: grids,
		},
		TableRows: trs,
	}
	f.Document.Body.Items = append(f.Document.Body.Items, tbl)
	return tbl
}

// Justification allows to set table's horizonal alignment
//
//	w:jc 属性的取值可以是以下之一：
//		start：左对齐。
//		center：居中对齐。
//		end：右对齐。
//		both：两端对齐。
//		distribute：分散对齐。
func (t *Table) Justification(val string) *Table {
	tp := t.ensureTableProperties()
	if tp.Justification == nil {
		tp.Justification = &Justification{Val: val}
		return t
	}
	tp.Justification.Val = val
	return t
}

// SetDefaultCellPadding sets table-level default cell padding (tblCellMar), unit: twips.
func (t *Table) SetDefaultCellPadding(top, right, bottom, left int64) *Table {
	tp := t.ensureTableProperties()
	if tp.CellMargins == nil {
		tp.CellMargins = &WTableDefaultCellMargins{}
	}
	tp.CellMargins.Top = &WTableCellMargin{W: top, Type: "dxa"}
	tp.CellMargins.Right = &WTableCellMargin{W: right, Type: "dxa"}
	tp.CellMargins.Bottom = &WTableCellMargin{W: bottom, Type: "dxa"}
	tp.CellMargins.Left = &WTableCellMargin{W: left, Type: "dxa"}
	return t
}

// SetLayoutFixed sets table layout to fixed.
func (t *Table) SetLayoutFixed() *Table {
	t.ensureTableProperties().Layout = &WTableLayout{Type: "fixed"}
	return t
}

// SetLayoutAutofit sets table layout to autofit.
func (t *Table) SetLayoutAutofit() *Table {
	t.ensureTableProperties().Layout = &WTableLayout{Type: "autofit"}
	return t
}

// SetWidthTwips sets table width in twips.
func (t *Table) SetWidthTwips(width int64) *Table {
	tp := t.ensureTableProperties()
	if width <= 0 {
		tp.Width = &WTableWidth{Type: "auto"}
		return t
	}
	tp.Width = &WTableWidth{W: width, Type: "dxa"}
	return t
}

// Justification allows to set table's horizonal alignment
//
//	w:jc 属性的取值可以是以下之一：
//		start：左对齐。
//		center：居中对齐。
//		end：右对齐。
//		both：两端对齐。
//		distribute：分散对齐。
func (w *WTableRow) Justification(val string) *WTableRow {
	if w.TableRowProperties.Justification == nil {
		w.TableRowProperties.Justification = &Justification{Val: val}
		return w
	}
	w.TableRowProperties.Justification.Val = val
	return w
}

// Shade allows to set cell's shade
func (c *WTableCell) Shade(val, color, fill string) *WTableCell {
	tcpr := c.ensureCellProperties()
	tcpr.Shade = &Shade{
		Val:   val,
		Color: color,
		Fill:  fill,
	}
	return c
}

// Padding allows to set cell's inner spacing (CSS-like top/right/bottom/left), unit: twips.
func (c *WTableCell) Padding(top, right, bottom, left int64) *WTableCell {
	tcpr := c.ensureCellProperties()
	if tcpr.Margins == nil {
		tcpr.Margins = &WTableCellMargins{}
	}
	tcpr.Margins.Top = &WTableCellMargin{W: top, Type: "dxa"}
	tcpr.Margins.Right = &WTableCellMargin{W: right, Type: "dxa"}
	tcpr.Margins.Bottom = &WTableCellMargin{W: bottom, Type: "dxa"}
	tcpr.Margins.Left = &WTableCellMargin{W: left, Type: "dxa"}
	return c
}

// SetColSpan sets the horizontal span for current cell.
// cols <= 1 clears gridSpan and restores default single-column behavior.
func (c *WTableCell) SetColSpan(cols int) *WTableCell {
	tcpr := c.ensureCellProperties()
	if cols <= 1 {
		tcpr.GridSpan = nil
		return c
	}
	tcpr.GridSpan = &WGridSpan{Val: cols}
	return c
}

// SetRowSpanRestart starts a vertical merge group at current cell.
func (c *WTableCell) SetRowSpanRestart() *WTableCell {
	c.ensureCellProperties().VMerge = &WvMerge{Val: "restart"}
	return c
}

// SetRowSpanContinue marks current cell as a continuation of vertical merge.
func (c *WTableCell) SetRowSpanContinue() *WTableCell {
	c.ensureCellProperties().VMerge = &WvMerge{}
	return c
}

// ClearRowSpan removes vertical merge setting from current cell.
func (c *WTableCell) ClearRowSpan() *WTableCell {
	c.ensureCellProperties().VMerge = nil
	return c
}

// SetCellBorderTop sets top border for current cell.
func (c *WTableCell) SetCellBorderTop(val string, size, space int, color string) *WTableCell {
	c.ensureCellBorders().Top = &WTableBorder{Val: val, Size: size, Space: space, Color: color}
	return c
}

// SetCellBorderRight sets right border for current cell.
func (c *WTableCell) SetCellBorderRight(val string, size, space int, color string) *WTableCell {
	c.ensureCellBorders().Right = &WTableBorder{Val: val, Size: size, Space: space, Color: color}
	return c
}

// SetCellBorderBottom sets bottom border for current cell.
func (c *WTableCell) SetCellBorderBottom(val string, size, space int, color string) *WTableCell {
	c.ensureCellBorders().Bottom = &WTableBorder{Val: val, Size: size, Space: space, Color: color}
	return c
}

// SetCellBorderLeft sets left border for current cell.
func (c *WTableCell) SetCellBorderLeft(val string, size, space int, color string) *WTableCell {
	c.ensureCellBorders().Left = &WTableBorder{Val: val, Size: size, Space: space, Color: color}
	return c
}

// SetCellBordersSame sets all four borders to the same style.
func (c *WTableCell) SetCellBordersSame(val string, size, space int, color string) *WTableCell {
	c.SetCellBorderTop(val, size, space, color)
	c.SetCellBorderRight(val, size, space, color)
	c.SetCellBorderBottom(val, size, space, color)
	c.SetCellBorderLeft(val, size, space, color)
	return c
}

// ClearCellBorders removes tcBorders from current cell.
func (c *WTableCell) ClearCellBorders() *WTableCell {
	c.ensureCellProperties().TableBorders = nil
	return c
}

func (c *WTableCell) ensureCellBorders() *WTableBorders {
	tcpr := c.ensureCellProperties()
	if tcpr.TableBorders == nil {
		tcpr.TableBorders = &WTableBorders{}
	}
	return tcpr.TableBorders
}

func (c *WTableCell) ensureCellProperties() *WTableCellProperties {
	if c.TableCellProperties == nil {
		c.TableCellProperties = &WTableCellProperties{}
		if len(c.ordered) > 0 {
			c.ordered = append([]interface{}{c.TableCellProperties}, c.ordered...)
		}
	}
	return c.TableCellProperties
}

func (t *Table) ensureTableProperties() *WTableProperties {
	if t.TableProperties == nil {
		t.TableProperties = &WTableProperties{}
		if len(t.ordered) > 0 {
			t.ordered = append([]interface{}{t.TableProperties}, t.ordered...)
		}
	}
	return t.TableProperties
}

// APITableBorderColors customizable param
type APITableBorderColors struct {
	Top     string
	Left    string
	Bottom  string
	Right   string
	InsideH string
	InsideV string
}

func (tbc *APITableBorderColors) applyDefault() {
	tbcR := reflect.ValueOf(tbc).Elem()

	for i := 0; i < tbcR.NumField(); i++ {
		if tbcR.Field(i).IsZero() {
			tbcR.Field(i).SetString("#000000")
		}
	}
}
