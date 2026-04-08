package docx

func (f *Docx) ensureMainSectPr(create bool) *SectPr {
	var main *SectPr
	for i := len(f.Document.Body.Items) - 1; i >= 0; i-- {
		if s, ok := f.Document.Body.Items[i].(*SectPr); ok {
			main = s
			break
		}
	}
	if main != nil || !create {
		return main
	}
	main = &SectPr{}
	f.Document.Body.Items = append(f.Document.Body.Items, main)
	return main
}

func (f *Docx) appendBodyItemBeforeTrailingSectPr(item interface{}) {
	items := f.Document.Body.Items
	n := len(items)
	if n > 0 {
		if _, ok := items[n-1].(*SectPr); ok {
			f.Document.Body.Items = append(items[:n-1], item, items[n-1])
			return
		}
	}
	f.Document.Body.Items = append(items, item)
}

func headerKindsInOrder() []HeaderKind {
	return []HeaderKind{HeaderDefault, HeaderFirst, HeaderEven}
}

func footerKindsInOrder() []FooterKind {
	return []FooterKind{FooterDefault, FooterFirst, FooterEven}
}

func (s *SectPr) setHeaderFooterRefs(hrefs []*HeaderReference, frefs []*FooterReference) {
	s.HeaderRefs = hrefs
	s.FooterRefs = frefs
	if len(s.ordered) == 0 {
		return
	}
	rest := make([]interface{}, 0, len(s.ordered))
	for _, item := range s.ordered {
		switch item.(type) {
		case *HeaderReference, *FooterReference:
			continue
		default:
			rest = append(rest, item)
		}
	}
	newOrdered := make([]interface{}, 0, len(hrefs)+len(frefs)+len(rest))
	for _, r := range hrefs {
		newOrdered = append(newOrdered, r)
	}
	for _, r := range frefs {
		newOrdered = append(newOrdered, r)
	}
	newOrdered = append(newOrdered, rest...)
	s.ordered = newOrdered
}
