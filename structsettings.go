package docx

import (
	"encoding/xml"
	"io"
	"strings"
)

// OnOff represents common OOXML on/off flag elements.
type OnOff struct {
	Val string `xml:"w:val,attr,omitempty"`
}

// Settings represents word/settings.xml content.
type Settings struct {
	XMLName xml.Name `xml:"w:settings,omitempty"`

	EvenAndOddHeaders *OnOff `xml:"w:evenAndOddHeaders,omitempty"`

	ordered []interface{}
	attrs   []xml.Attr
}

// UnmarshalXML parses settings while preserving unknown nodes order.
func (s *Settings) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	s.XMLName = start.Name
	s.attrs = append([]xml.Attr(nil), start.Attr...)
	for {
		tok, err := d.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}
		tt, ok := tok.(xml.StartElement)
		if !ok {
			continue
		}
		switch tt.Name.Local {
		case "evenAndOddHeaders":
			var value OnOff
			err = d.DecodeElement(&value, &tt)
			if err != nil && !strings.HasPrefix(err.Error(), "expected") {
				return err
			}
			s.EvenAndOddHeaders = &value
			s.ordered = append(s.ordered, s.EvenAndOddHeaders)
		default:
			raw, err := decodeRawXMLNode(d, tt)
			if err != nil {
				return err
			}
			s.ordered = append(s.ordered, raw)
		}
	}
	return nil
}

// MarshalXML writes settings with stable children order.
func (s *Settings) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	start.Name = xml.Name{Local: "w:settings"}
	start.Attr = settingsRootAttrs(s.attrs)
	if err := e.EncodeToken(start); err != nil {
		return err
	}
	if len(s.ordered) > 0 {
		for _, item := range s.ordered {
			if item == nil {
				continue
			}
			switch v := item.(type) {
			case *OnOff:
				if v == s.EvenAndOddHeaders {
					if err := e.EncodeElement(v, xml.StartElement{Name: xml.Name{Local: "w:evenAndOddHeaders"}}); err != nil {
						return err
					}
					continue
				}
				if err := e.Encode(v); err != nil {
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
	if s.EvenAndOddHeaders != nil {
		if err := e.EncodeElement(s.EvenAndOddHeaders, xml.StartElement{Name: xml.Name{Local: "w:evenAndOddHeaders"}}); err != nil {
			return err
		}
	}
	return e.EncodeToken(start.End())
}

func settingsRootAttrs(existing []xml.Attr) []xml.Attr {
	attrs := make([]xml.Attr, 0, len(existing)+1)
	for _, attr := range existing {
		attrs = append(attrs, normalizeRootAttr(attr))
	}
	attrs = ensureXMLNS(attrs, "w", XMLNS_W)
	return attrs
}

func (s *Settings) setEvenAndOddHeaders(enabled bool) {
	if enabled {
		if s.EvenAndOddHeaders == nil {
			s.EvenAndOddHeaders = &OnOff{}
			if len(s.ordered) > 0 {
				s.ordered = append([]interface{}{s.EvenAndOddHeaders}, s.ordered...)
			}
		}
		return
	}
	if s.EvenAndOddHeaders == nil {
		return
	}
	toRemove := s.EvenAndOddHeaders
	s.EvenAndOddHeaders = nil
	if len(s.ordered) == 0 {
		return
	}
	next := make([]interface{}, 0, len(s.ordered))
	for _, item := range s.ordered {
		if item == toRemove {
			continue
		}
		next = append(next, item)
	}
	s.ordered = next
}
