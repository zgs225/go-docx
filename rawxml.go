package docx

import (
	"bytes"
	"encoding/xml"
	"io"
	"strings"
)

// RawXMLNode preserves an unsupported element for round-trip writeback.
//
// Phase 0 only guarantees lossless preservation and writeback order.
// It does not provide semantic editing for these nodes.
type RawXMLNode struct {
	Name     xml.Name
	Attrs    []xml.Attr
	OuterXML string
}

func decodeRawXMLNode(d *xml.Decoder, start xml.StartElement) (*RawXMLNode, error) {
	var buf bytes.Buffer
	enc := xml.NewEncoder(&buf)
	if err := enc.EncodeToken(start); err != nil {
		return nil, err
	}

	depth := 1
	for depth > 0 {
		tok, err := d.Token()
		if err != nil {
			return nil, err
		}
		switch tok.(type) {
		case xml.StartElement:
			depth++
		case xml.EndElement:
			depth--
		}
		if err := enc.EncodeToken(tok); err != nil {
			return nil, err
		}
	}
	if err := enc.Flush(); err != nil {
		return nil, err
	}
	return &RawXMLNode{
		Name:     start.Name,
		Attrs:    start.Attr,
		OuterXML: buf.String(),
	}, nil
}

// MarshalXML writes preserved XML content back without semantic conversion.
func (r *RawXMLNode) MarshalXML(e *xml.Encoder, _ xml.StartElement) error {
	if r == nil {
		return nil
	}
	if r.OuterXML == "" {
		start := xml.StartElement{Name: r.Name, Attr: r.Attrs}
		if err := e.EncodeToken(start); err != nil {
			return err
		}
		return e.EncodeToken(start.End())
	}

	dec := xml.NewDecoder(strings.NewReader(r.OuterXML))
	for {
		tok, err := dec.Token()
		if err == io.EOF {
			return nil
		}
		if err != nil {
			return err
		}
		if err := e.EncodeToken(tok); err != nil {
			return err
		}
	}
}
