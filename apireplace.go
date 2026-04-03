package docx

import (
	"errors"
	"sort"
	"strings"
)

var errEmptyFindText = errors.New("replace find text cannot be empty")

// ReplaceOption customizes text replacement behavior.
type ReplaceOption func(*replaceConfig)

type replaceConfig struct {
	maxReplacements int
	caseSensitive   bool
}

func defaultReplaceConfig() replaceConfig {
	return replaceConfig{
		maxReplacements: -1,
		caseSensitive:   true,
	}
}

// WithMaxReplacements limits total replacement count.
// n < 0 means unlimited, n == 0 means no replacements.
func WithMaxReplacements(n int) ReplaceOption {
	return func(c *replaceConfig) {
		c.maxReplacements = n
	}
}

// WithCaseSensitive sets whether find matches are case-sensitive.
func WithCaseSensitive(enabled bool) ReplaceOption {
	return func(c *replaceConfig) {
		c.caseSensitive = enabled
	}
}

// ReplaceText replaces all matches of find with replace across paragraphs and table cell paragraphs.
// It supports matches spanning contiguous text-only runs.
func (f *Docx) ReplaceText(find, replace string, opts ...ReplaceOption) error {
	if find == "" {
		return errEmptyFindText
	}
	cfg := defaultReplaceConfig()
	for _, opt := range opts {
		if opt != nil {
			opt(&cfg)
		}
	}
	if cfg.maxReplacements == 0 {
		return nil
	}
	remaining := cfg.maxReplacements
	for _, item := range f.Document.Body.Items {
		if remaining == 0 {
			break
		}
		switch v := item.(type) {
		case *Paragraph:
			remaining = replaceInParagraph(v, find, replace, cfg, remaining)
		case *Table:
			remaining = replaceInTable(v, find, replace, cfg, remaining)
		}
	}
	return nil
}

// ReplacePlaceholder replaces placeholders in format {{key}}.
// Missing keys are ignored.
func (f *Docx) ReplacePlaceholder(data map[string]string, opts ...ReplaceOption) error {
	if len(data) == 0 {
		return nil
	}
	keys := make([]string, 0, len(data))
	for k := range data {
		keys = append(keys, k)
	}
	sort.Strings(keys)

	cfg := defaultReplaceConfig()
	for _, opt := range opts {
		if opt != nil {
			opt(&cfg)
		}
	}
	remaining := cfg.maxReplacements
	for _, k := range keys {
		if remaining == 0 {
			break
		}
		placeholder := "{{" + k + "}}"
		remaining = replaceAcrossDoc(f, placeholder, data[k], cfg, remaining)
	}
	return nil
}

func replaceAcrossDoc(f *Docx, find, replace string, cfg replaceConfig, remaining int) int {
	for _, item := range f.Document.Body.Items {
		if remaining == 0 {
			break
		}
		switch v := item.(type) {
		case *Paragraph:
			remaining = replaceInParagraph(v, find, replace, cfg, remaining)
		case *Table:
			remaining = replaceInTable(v, find, replace, cfg, remaining)
		}
	}
	return remaining
}

func replaceInTable(t *Table, find, replace string, cfg replaceConfig, remaining int) int {
	for _, row := range t.TableRows {
		for _, cell := range row.TableCells {
			for _, p := range cell.Paragraphs {
				if remaining == 0 {
					return remaining
				}
				remaining = replaceInParagraph(p, find, replace, cfg, remaining)
			}
			for _, nt := range cell.Tables {
				if remaining == 0 {
					return remaining
				}
				remaining = replaceInTable(nt, find, replace, cfg, remaining)
			}
		}
	}
	return remaining
}

func replaceInParagraph(p *Paragraph, find, replace string, cfg replaceConfig, remaining int) int {
	runGroups := paragraphTextRunGroups(p)
	if len(runGroups) == 0 {
		return remaining
	}
	for _, group := range runGroups {
		if remaining == 0 {
			break
		}
		used := replaceInRunGroup(group, find, replace, cfg, remaining)
		if remaining > 0 {
			remaining -= used
		}
	}
	return remaining
}

func paragraphTextRunGroups(p *Paragraph) [][]*Run {
	if len(p.ordered) > 0 {
		return collectRunGroupsFromItems(p.ordered)
	}
	return collectRunGroupsFromItems(p.Children)
}

func collectRunGroupsFromItems(items []interface{}) [][]*Run {
	groups := make([][]*Run, 0, 8)
	cur := make([]*Run, 0, 8)
	flush := func() {
		if len(cur) > 0 {
			groups = append(groups, cur)
			cur = make([]*Run, 0, 8)
		}
	}
	for _, item := range items {
		r, ok := item.(*Run)
		if !ok {
			flush()
			continue
		}
		if runText, ok := runPlainText(r); ok && runText != "" {
			cur = append(cur, r)
			continue
		}
		flush()
	}
	flush()
	return groups
}

func runPlainText(r *Run) (string, bool) {
	if r == nil || len(r.Children) == 0 {
		return "", false
	}
	var sb strings.Builder
	for _, child := range r.Children {
		txt, ok := child.(*Text)
		if !ok {
			return "", false
		}
		sb.WriteString(txt.Text)
	}
	return sb.String(), true
}

type textMatch struct {
	Start int
	End   int
}

func replaceInRunGroup(runs []*Run, find, replace string, cfg replaceConfig, remaining int) int {
	if len(runs) == 0 {
		return 0
	}
	texts := make([]string, 0, len(runs))
	var total strings.Builder
	offsets := make([]int, 0, len(runs))
	pos := 0
	for _, r := range runs {
		t, ok := runPlainText(r)
		if !ok || t == "" {
			return 0
		}
		offsets = append(offsets, pos)
		texts = append(texts, t)
		total.WriteString(t)
		pos += len(t)
	}
	full := total.String()
	matches := findMatches(full, find, cfg.caseSensitive, remaining)
	if len(matches) == 0 {
		return 0
	}

	newTexts := make([]string, len(texts))
	for i := range texts {
		rs := offsets[i]
		re := rs + len(texts[i])
		cur := rs
		var sb strings.Builder
		for _, m := range matches {
			if m.End <= rs || m.Start >= re {
				continue
			}
			if m.Start > cur {
				sb.WriteString(full[cur:m.Start])
			}
			if m.Start >= rs && m.Start < re {
				sb.WriteString(replace)
			}
			if m.End > cur {
				cur = m.End
			}
		}
		if cur < re {
			sb.WriteString(full[cur:re])
		}
		newTexts[i] = sb.String()
	}

	for i, r := range runs {
		r.ordered = nil
		if newTexts[i] == "" {
			r.Children = nil
			continue
		}
		txt := &Text{Text: newTexts[i]}
		if strings.HasPrefix(newTexts[i], " ") || strings.HasSuffix(newTexts[i], " ") {
			txt.XMLSpace = "preserve"
		}
		r.Children = []interface{}{txt}
	}
	return len(matches)
}

func findMatches(source, find string, caseSensitive bool, remaining int) []textMatch {
	if find == "" || source == "" {
		return nil
	}
	src := source
	target := find
	if !caseSensitive {
		src = strings.ToLower(source)
		target = strings.ToLower(find)
	}

	matches := make([]textMatch, 0, 8)
	start := 0
	for {
		if remaining == 0 {
			break
		}
		idx := strings.Index(src[start:], target)
		if idx < 0 {
			break
		}
		s := start + idx
		e := s + len(target)
		matches = append(matches, textMatch{Start: s, End: e})
		start = e
		if remaining > 0 {
			remaining--
		}
	}
	return matches
}
