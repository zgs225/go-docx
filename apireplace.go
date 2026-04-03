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
	maxReplacements        int
	caseSensitive          bool
	enableFieldCodeReplace bool
	allowedFieldTypes      map[string]struct{}
}

func defaultReplaceConfig() replaceConfig {
	return replaceConfig{
		maxReplacements:        -1,
		caseSensitive:          true,
		enableFieldCodeReplace: false,
		allowedFieldTypes:      defaultFieldTypeWhitelist(),
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

// WithFieldCodeReplacement enables replacement for field code text (instrText).
func WithFieldCodeReplacement(enabled bool) ReplaceOption {
	return func(c *replaceConfig) {
		c.enableFieldCodeReplace = enabled
	}
}

// WithFieldCodeWhitelist overrides allowed field types for instrText replacement.
// Empty input resets to default whitelist.
func WithFieldCodeWhitelist(types ...string) ReplaceOption {
	return func(c *replaceConfig) {
		if len(types) == 0 {
			c.allowedFieldTypes = defaultFieldTypeWhitelist()
			return
		}
		wl := make(map[string]struct{}, len(types))
		for _, t := range types {
			t = strings.ToUpper(strings.TrimSpace(t))
			if t == "" {
				continue
			}
			wl[t] = struct{}{}
		}
		if len(wl) == 0 {
			c.allowedFieldTypes = defaultFieldTypeWhitelist()
			return
		}
		c.allowedFieldTypes = wl
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
	if len(runGroups) > 0 {
		for _, group := range runGroups {
			if remaining == 0 {
				break
			}
			used := replaceInRunGroup(group, find, replace, cfg, remaining)
			if remaining > 0 {
				remaining -= used
			}
		}
	}
	if cfg.enableFieldCodeReplace && remaining != 0 {
		used := replaceInFieldCodes(p, find, replace, cfg, remaining)
		if remaining > 0 {
			remaining -= used
		}
	}
	cleanupParagraphAfterReplace(p)
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
		switch v := item.(type) {
		case *Run:
			if runText, ok := runPlainText(v); ok && runText != "" {
				cur = append(cur, v)
				continue
			}
			flush()
		case *Hyperlink:
			flush() // do not cross boundaries with non-hyperlink runs
			if runText, ok := runPlainText(&v.Run); ok && runText != "" {
				groups = append(groups, []*Run{&v.Run})
			}
		default:
			flush()
		}
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

func cleanupParagraphAfterReplace(p *Paragraph) {
	if len(p.ordered) > 0 {
		newOrdered := make([]interface{}, 0, len(p.ordered))
		newChildren := make([]interface{}, 0, len(p.Children))
		for _, item := range p.ordered {
			switch v := item.(type) {
			case *Run:
				if !runShouldRemain(v) {
					continue
				}
				newOrdered = append(newOrdered, v)
				newChildren = append(newChildren, v)
			case *Hyperlink:
				if !runShouldRemain(&v.Run) {
					continue
				}
				newOrdered = append(newOrdered, v)
				newChildren = append(newChildren, v)
			case *RawXMLNode:
				newOrdered = append(newOrdered, v)
			default:
				newOrdered = append(newOrdered, item)
				newChildren = append(newChildren, item)
			}
		}
		p.ordered = newOrdered
		p.Children = newChildren
		return
	}

	newChildren := make([]interface{}, 0, len(p.Children))
	for _, item := range p.Children {
		switch v := item.(type) {
		case *Run:
			if !runShouldRemain(v) {
				continue
			}
			newChildren = append(newChildren, v)
		case *Hyperlink:
			if !runShouldRemain(&v.Run) {
				continue
			}
			newChildren = append(newChildren, v)
		default:
			newChildren = append(newChildren, item)
		}
	}
	p.Children = newChildren
}

func runShouldRemain(r *Run) bool {
	if len(r.Children) > 0 {
		return true
	}
	if r.InstrText != "" {
		return true
	}
	return len(r.ordered) > 0
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

func defaultFieldTypeWhitelist() map[string]struct{} {
	return map[string]struct{}{
		"FORMTEXT":   {},
		"MERGEFIELD": {},
	}
}

func replaceInFieldCodes(p *Paragraph, find, replace string, cfg replaceConfig, remaining int) int {
	fields := collectParagraphFields(p)
	if len(fields) == 0 {
		return 0
	}
	used := 0
	for _, f := range fields {
		if remaining == 0 {
			break
		}
		ft := detectFieldType(f.codeRuns)
		if ft == "" {
			continue
		}
		if _, ok := cfg.allowedFieldTypes[ft]; !ok {
			continue
		}
		n := replaceInInstrRuns(f.codeRuns, find, replace, cfg.caseSensitive, remaining)
		if n == 0 {
			continue
		}
		used += n
		if remaining > 0 {
			remaining -= n
		}
	}
	return used
}

type paragraphField struct {
	codeRuns []*Run
}

func collectParagraphFields(p *Paragraph) []paragraphField {
	refs := collectRunRefs(p)
	fields := make([]paragraphField, 0, 8)
	active := false
	start := -1
	separate := -1

	for i, ref := range refs {
		if ref.run == nil {
			continue
		}
		if runHasFldCharType(ref.run, "begin") {
			active = true
			start = i
			separate = -1
			continue
		}
		if !active {
			continue
		}
		if runHasFldCharType(ref.run, "separate") && separate == -1 {
			separate = i
		}
		if runHasFldCharType(ref.run, "end") {
			endCode := i
			if separate != -1 {
				endCode = separate
			}
			codeRuns := make([]*Run, 0, endCode-start+1)
			for j := start; j <= endCode; j++ {
				if refs[j].run != nil {
					codeRuns = append(codeRuns, refs[j].run)
				}
			}
			fields = append(fields, paragraphField{codeRuns: codeRuns})
			active = false
			start = -1
			separate = -1
		}
	}
	return fields
}

type runRef struct {
	run *Run
}

func collectRunRefs(p *Paragraph) []runRef {
	items := p.Children
	if len(p.ordered) > 0 {
		items = p.ordered
	}
	refs := make([]runRef, 0, len(items))
	for _, item := range items {
		switch v := item.(type) {
		case *Run:
			refs = append(refs, runRef{run: v})
		case *Hyperlink:
			refs = append(refs, runRef{run: &v.Run})
		}
	}
	return refs
}

func runHasFldCharType(r *Run, want string) bool {
	items := r.Children
	if len(r.ordered) > 0 {
		items = r.ordered
	}
	for _, item := range items {
		raw, ok := item.(*RawXMLNode)
		if !ok {
			continue
		}
		if raw.Name.Local != "fldChar" {
			continue
		}
		val := getAtt(raw.Attrs, "fldCharType")
		if strings.EqualFold(val, want) {
			return true
		}
	}
	return false
}

func detectFieldType(runs []*Run) string {
	var sb strings.Builder
	for _, r := range runs {
		if r == nil || r.InstrText == "" {
			continue
		}
		sb.WriteString(r.InstrText)
	}
	parts := strings.Fields(sb.String())
	if len(parts) == 0 {
		return ""
	}
	return strings.ToUpper(parts[0])
}

func replaceInInstrRuns(runs []*Run, find, replace string, caseSensitive bool, remaining int) int {
	type pos struct {
		run *Run
		s   int
		e   int
	}
	positions := make([]pos, 0, len(runs))
	var sb strings.Builder
	cur := 0
	for _, r := range runs {
		if r == nil || r.InstrText == "" {
			continue
		}
		s := cur
		sb.WriteString(r.InstrText)
		cur += len(r.InstrText)
		positions = append(positions, pos{run: r, s: s, e: cur})
	}
	if len(positions) == 0 {
		return 0
	}

	full := sb.String()
	matches := findMatches(full, find, caseSensitive, remaining)
	if len(matches) == 0 {
		return 0
	}

	for _, p := range positions {
		cur = p.s
		var out strings.Builder
		for _, m := range matches {
			if m.End <= p.s || m.Start >= p.e {
				continue
			}
			if m.Start > cur {
				out.WriteString(full[cur:m.Start])
			}
			if m.Start >= p.s && m.Start < p.e {
				out.WriteString(replace)
			}
			if m.End > cur {
				cur = m.End
			}
		}
		if cur < p.e {
			out.WriteString(full[cur:p.e])
		}
		setRunInstrText(p.run, out.String())
	}
	return len(matches)
}

func setRunInstrText(r *Run, text string) {
	r.InstrText = text
	if len(r.ordered) == 0 {
		return
	}
	found := false
	for _, item := range r.ordered {
		if it, ok := item.(*runInstrText); ok {
			it.Text = text
			found = true
			break
		}
	}
	if !found && text != "" {
		r.ordered = append(r.ordered, &runInstrText{Text: text})
	}
}
