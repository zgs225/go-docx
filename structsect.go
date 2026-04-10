/*
   Copyright (c) 2024 mabiao0525 (马飚)

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
	"encoding/xml"
	"io"
	"strconv"
	"strings"
)

// SectPr show the properties of the document, like paper size
type SectPr struct {
	XMLName    xml.Name           `xml:"w:sectPr,omitempty"` // properties of the document, including paper size
	HeaderRefs []*HeaderReference `xml:"w:headerReference,omitempty"`
	FooterRefs []*FooterReference `xml:"w:footerReference,omitempty"`
	TitlePg    *OnOff             `xml:"w:titlePg,omitempty"`
	PgSz       *PgSz              `xml:"w:pgSz,omitempty"`
	PgNumType  *PgNumType         `xml:"w:pgNumType,omitempty"`
	PgMar      *PgMar             `xml:"w:pgMar,omitempty"`
	Cols       *Cols              `xml:"w:cols,omitempty"`
	DocGrid    *DocGrid           `xml:"w:docGrid,omitempty"`
	ordered    []interface{}      // keeps known/unknown order for round-trip
}

// HeaderReference links a section to a header relationship.
type HeaderReference struct {
	XMLName xml.Name `xml:"w:headerReference,omitempty"`
	Type    string   `xml:"w:type,attr,omitempty"`
	RID     string   `xml:"r:id,attr,omitempty"`
}

// FooterReference links a section to a footer relationship.
type FooterReference struct {
	XMLName xml.Name `xml:"w:footerReference,omitempty"`
	Type    string   `xml:"w:type,attr,omitempty"`
	RID     string   `xml:"r:id,attr,omitempty"`
}

// PgSz show the paper size
type PgSz struct {
	W int `xml:"w:w,attr"` // width of paper
	H int `xml:"w:h,attr"` // high of paper
}

// PgNumType defines page-number formatting for a section.
type PgNumType struct {
	Fmt string `xml:"w:fmt,attr,omitempty"`
}

// PgMar show the page margin
type PgMar struct {
	Top    int `xml:"w:top,attr"`
	Left   int `xml:"w:left,attr"`
	Bottom int `xml:"w:bottom,attr"`
	Right  int `xml:"w:right,attr"`
	Header int `xml:"w:header,attr"`
	Footer int `xml:"w:footer,attr"`
	Gutter int `xml:"w:gutter,attr"`
}

// Cols show the number of columns
type Cols struct {
	Space int `xml:"w:space,attr"`
}

// DocGrid show the document grid
type DocGrid struct {
	Type      string `xml:"w:type,attr"`
	LinePitch int    `xml:"w:linePitch,attr"`
}

// UnmarshalXML ...
func (sect *SectPr) UnmarshalXML(d *xml.Decoder, _ xml.StartElement) error {
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}
		if tt, ok := t.(xml.StartElement); ok {
			switch tt.Name.Local {
			case "headerReference":
				ref := &HeaderReference{
					Type: getAtt(tt.Attr, "type"),
					RID:  getAtt(tt.Attr, "id"),
				}
				sect.HeaderRefs = append(sect.HeaderRefs, ref)
				sect.ordered = append(sect.ordered, ref)
				err = d.Skip()
				if err != nil {
					return err
				}
			case "footerReference":
				ref := &FooterReference{
					Type: getAtt(tt.Attr, "type"),
					RID:  getAtt(tt.Attr, "id"),
				}
				sect.FooterRefs = append(sect.FooterRefs, ref)
				sect.ordered = append(sect.ordered, ref)
				err = d.Skip()
				if err != nil {
					return err
				}
			case "pgSz":
				var value PgSz
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				sect.PgSz = &value
				sect.ordered = append(sect.ordered, sect.PgSz)
			case "titlePg":
				var value OnOff
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				sect.TitlePg = &value
				sect.ordered = append(sect.ordered, sect.TitlePg)
			case "pgNumType":
				var value PgNumType
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				sect.PgNumType = &value
				sect.ordered = append(sect.ordered, sect.PgNumType)
			case "pgMar":
				var value PgMar
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				sect.PgMar = &value
				sect.ordered = append(sect.ordered, sect.PgMar)
			case "cols":
				var value Cols
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				sect.Cols = &value
				sect.ordered = append(sect.ordered, sect.Cols)
			case "docGrid":
				var value DocGrid
				err = d.DecodeElement(&value, &tt)
				if err != nil && !strings.HasPrefix(err.Error(), "expected") {
					return err
				}
				sect.DocGrid = &value
				sect.ordered = append(sect.ordered, sect.DocGrid)
			default:
				raw, err := decodeRawXMLNode(d, tt)
				if err != nil {
					return err
				}
				sect.ordered = append(sect.ordered, raw)
			}
		}
	}
	return nil
}

// MarshalXML keeps section children write order stable for round-trip.
func (sect *SectPr) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = xml.Name{Local: "w:sectPr"}
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if len(sect.ordered) > 0 {
		for _, item := range sect.ordered {
			if item == nil {
				continue
			}
			switch v := item.(type) {
			case *PgSz:
				if err := e.EncodeElement(v, xml.StartElement{Name: xml.Name{Local: "w:pgSz"}}); err != nil {
					return err
				}
			case *OnOff:
				if v == sect.TitlePg {
					if err := e.EncodeElement(v, xml.StartElement{Name: xml.Name{Local: "w:titlePg"}}); err != nil {
						return err
					}
					continue
				}
				if err := e.Encode(v); err != nil {
					return err
				}
			case *PgNumType:
				if err := e.EncodeElement(v, xml.StartElement{Name: xml.Name{Local: "w:pgNumType"}}); err != nil {
					return err
				}
			case *PgMar:
				if err := e.EncodeElement(v, xml.StartElement{Name: xml.Name{Local: "w:pgMar"}}); err != nil {
					return err
				}
			case *Cols:
				if err := e.EncodeElement(v, xml.StartElement{Name: xml.Name{Local: "w:cols"}}); err != nil {
					return err
				}
			case *DocGrid:
				if err := e.EncodeElement(v, xml.StartElement{Name: xml.Name{Local: "w:docGrid"}}); err != nil {
					return err
				}
			default:
				if err := e.Encode(item); err != nil {
					return err
				}
			}
		}
		return e.EncodeToken(start.End())
	}
	for _, ref := range sect.HeaderRefs {
		if ref == nil {
			continue
		}
		if err := e.EncodeElement(ref, xml.StartElement{Name: xml.Name{Local: "w:headerReference"}}); err != nil {
			return err
		}
	}
	for _, ref := range sect.FooterRefs {
		if ref == nil {
			continue
		}
		if err := e.EncodeElement(ref, xml.StartElement{Name: xml.Name{Local: "w:footerReference"}}); err != nil {
			return err
		}
	}
	if sect.TitlePg != nil {
		if err := e.EncodeElement(sect.TitlePg, xml.StartElement{Name: xml.Name{Local: "w:titlePg"}}); err != nil {
			return err
		}
	}
	if sect.PgSz != nil {
		if err := e.EncodeElement(sect.PgSz, xml.StartElement{Name: xml.Name{Local: "w:pgSz"}}); err != nil {
			return err
		}
	}
	if sect.PgNumType != nil {
		if err := e.EncodeElement(sect.PgNumType, xml.StartElement{Name: xml.Name{Local: "w:pgNumType"}}); err != nil {
			return err
		}
	}
	if sect.PgMar != nil {
		if err := e.EncodeElement(sect.PgMar, xml.StartElement{Name: xml.Name{Local: "w:pgMar"}}); err != nil {
			return err
		}
	}
	if sect.Cols != nil {
		if err := e.EncodeElement(sect.Cols, xml.StartElement{Name: xml.Name{Local: "w:cols"}}); err != nil {
			return err
		}
	}
	if sect.DocGrid != nil {
		if err := e.EncodeElement(sect.DocGrid, xml.StartElement{Name: xml.Name{Local: "w:docGrid"}}); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

// UnmarshalXML ...
func (pgsz *PgSz) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	var err error

	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "w":
			pgsz.W, err = strconv.Atoi(attr.Value)
			if err != nil {
				return err
			}
		case "h":
			pgsz.H, err = strconv.Atoi(attr.Value)
			if err != nil {
				return err
			}
		default:
			// ignore other attributes now
		}
	}
	// Consume the end element
	_, err = d.Token()
	return err
}

// UnmarshalXML ...
func (p *PgNumType) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "fmt":
			p.Fmt = attr.Value
		default:
			// ignore other attributes now
		}
	}
	// Consume the end element.
	_, err := d.Token()
	return err
}

// UnmarshalXML ...
func (pgmar *PgMar) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	var err error

	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "top":
			pgmar.Top, err = strconv.Atoi(attr.Value)
			if err != nil {
				return err
			}
		case "left":
			pgmar.Left, err = strconv.Atoi(attr.Value)
			if err != nil {
				return err
			}
		case "bottom":
			pgmar.Bottom, err = strconv.Atoi(attr.Value)
			if err != nil {
				return err
			}
		case "right":
			pgmar.Right, err = strconv.Atoi(attr.Value)
			if err != nil {
				return err
			}
		case "header":
			pgmar.Header, err = strconv.Atoi(attr.Value)
			if err != nil {
				return err
			}
		case "footer":
			pgmar.Footer, err = strconv.Atoi(attr.Value)
			if err != nil {
				return err
			}
		case "gutter":
			pgmar.Gutter, err = strconv.Atoi(attr.Value)
			if err != nil {
				return err
			}
		default:
			// ignore other attributes now
		}
	}
	// Consume the end element
	_, err = d.Token()
	return err
}

// UnmarshalXML ...
func (cols *Cols) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	var err error

	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "space":
			cols.Space, err = strconv.Atoi(attr.Value)
			if err != nil {
				return err
			}
		default:
			// ignore other attributes now
		}
	}
	// Consume the end element
	_, err = d.Token()
	return err
}

// UnmarshalXML ...
func (dg *DocGrid) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	var err error

	for _, attr := range start.Attr {
		switch attr.Name.Local {
		case "linePitch":
			dg.LinePitch, err = strconv.Atoi(attr.Value)
			if err != nil {
				return err
			}
		case "type":
			dg.Type = attr.Value
		default:
			// ignore other attributes now
		}
	}
	// Consume the end element
	_, err = d.Token()
	return err
}
