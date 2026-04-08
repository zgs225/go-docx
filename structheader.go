package docx

import (
	"encoding/xml"
	"io"
	"strings"
)

// HeaderKind identifies header reference type in sectPr.
type HeaderKind string

const (
	HeaderDefault HeaderKind = "default"
	HeaderFirst   HeaderKind = "first"
	HeaderEven    HeaderKind = "even"
)

// FooterKind identifies footer reference type in sectPr.
type FooterKind string

const (
	FooterDefault FooterKind = "default"
	FooterFirst   FooterKind = "first"
	FooterEven    FooterKind = "even"
)

// PageNumberStyle defines PAGE field formatting switch.
type PageNumberStyle string

const (
	PageNumberArabic      PageNumberStyle = "arabic"
	PageNumberRoman       PageNumberStyle = "roman"
	PageNumberRomanUpper  PageNumberStyle = "ROMAN"
	PageNumberLetter      PageNumberStyle = "letter"
	PageNumberLetterUpper PageNumberStyle = "LETTER"
)

// Header represents word/header*.xml content.
type Header struct {
	XMLName xml.Name `xml:"w:hdr,omitempty"`
	Items   []interface{}
	ordered []interface{}
	attrs   []xml.Attr

	file *Docx
}

// Footer represents word/footer*.xml content.
type Footer struct {
	XMLName xml.Name `xml:"w:ftr,omitempty"`
	Items   []interface{}
	ordered []interface{}
	attrs   []xml.Attr

	file *Docx
}

// UnmarshalXML ...
func (h *Header) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	h.XMLName = start.Name
	h.attrs = append([]xml.Attr(nil), start.Attr...)
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}
		tt, ok := t.(xml.StartElement)
		if !ok {
			continue
		}
		item, err := decodeHeaderFooterItem(d, tt, h.file)
		if err != nil {
			return err
		}
		h.Items = append(h.Items, item)
		h.ordered = append(h.ordered, item)
	}
	return nil
}

// MarshalXML keeps child order stable for round-trip.
func (h *Header) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = xml.Name{Local: "w:hdr"}
	start.Attr = headerFooterRootAttrs(h.attrs)
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	items := h.Items
	if len(h.ordered) > 0 {
		items = h.ordered
	}
	for _, item := range items {
		if item == nil {
			continue
		}
		if err := e.Encode(item); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

// UnmarshalXML ...
func (f *Footer) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	f.XMLName = start.Name
	f.attrs = append([]xml.Attr(nil), start.Attr...)
	for {
		t, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}
		tt, ok := t.(xml.StartElement)
		if !ok {
			continue
		}
		item, err := decodeHeaderFooterItem(d, tt, f.file)
		if err != nil {
			return err
		}
		f.Items = append(f.Items, item)
		f.ordered = append(f.ordered, item)
	}
	return nil
}

// MarshalXML keeps child order stable for round-trip.
func (f *Footer) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = xml.Name{Local: "w:ftr"}
	start.Attr = headerFooterRootAttrs(f.attrs)
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	items := f.Items
	if len(f.ordered) > 0 {
		items = f.ordered
	}
	for _, item := range items {
		if item == nil {
			continue
		}
		if err := e.Encode(item); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

func headerFooterRootAttrs(existing []xml.Attr) []xml.Attr {
	attrs := make([]xml.Attr, 0, len(existing)+9)
	for _, attr := range existing {
		attrs = append(attrs, normalizeRootAttr(attr))
	}
	attrs = ensureXMLNS(attrs, "w", XMLNS_W)
	attrs = ensureXMLNS(attrs, "r", XMLNS_R)
	attrs = ensureXMLNS(attrs, "wp", XMLNS_WP)
	attrs = ensureXMLNS(attrs, "wps", XMLNS_WPS)
	attrs = ensureXMLNS(attrs, "wpc", XMLNS_WPC)
	attrs = ensureXMLNS(attrs, "wpg", XMLNS_WPG)
	attrs = ensureXMLNS(attrs, "mc", XMLNS_MC)
	attrs = ensureXMLNS(attrs, "o", XMLNS_O)
	attrs = ensureXMLNS(attrs, "v", XMLNS_V)
	return attrs
}

func ensureXMLNS(attrs []xml.Attr, prefix, uri string) []xml.Attr {
	for _, attr := range attrs {
		if attr.Name.Space == "xmlns" && attr.Name.Local == prefix {
			return attrs
		}
		if attr.Name.Space == "" && attr.Name.Local == "xmlns:"+prefix {
			return attrs
		}
	}
	return append(attrs, xml.Attr{
		Name:  xml.Name{Local: "xmlns:" + prefix},
		Value: uri,
	})
}

func normalizeRootAttr(attr xml.Attr) xml.Attr {
	if attr.Name.Space == "xmlns" {
		attr.Name = xml.Name{Local: "xmlns:" + attr.Name.Local}
	}
	return attr
}

func decodeHeaderFooterItem(d *xml.Decoder, tt xml.StartElement, file *Docx) (interface{}, error) {
	switch tt.Name.Local {
	case "p":
		var value Paragraph
		value.file = file
		err := d.DecodeElement(&value, &tt)
		if err != nil && !strings.HasPrefix(err.Error(), "expected") {
			return nil, err
		}
		return &value, nil
	case "tbl":
		var value Table
		value.file = file
		err := d.DecodeElement(&value, &tt)
		if err != nil && !strings.HasPrefix(err.Error(), "expected") {
			return nil, err
		}
		return &value, nil
	default:
		return decodeRawXMLNode(d, tt)
	}
}
