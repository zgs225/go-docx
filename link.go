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
	"errors"
	"strconv"
	"strings"
	"sync/atomic"
)

var (
	// ErrRefIDNotFound cannot find such reference
	ErrRefIDNotFound = errors.New("ref id not found")
	// ErrRefTargetNotFound cannot find such target
	ErrRefTargetNotFound = errors.New("ref target not found")
)

// when adding an hyperlink we need to store a reference in the relationship field
//
//	this func is not thread-safe
func (f *Docx) addLinkRelation(link string) string {
	rel := Relationship{
		ID:         "rId" + strconv.Itoa(int(atomic.AddUintptr(&f.rID, 1))),
		Type:       REL_HYPERLINK,
		Target:     link,
		TargetMode: REL_TARGETMODE,
	}

	f.docRelation.Relationship = append(f.docRelation.Relationship, rel)

	return rel.ID
}

// when adding an image we need to store a reference in the relationship field
//
//	this func is not thread-safe
func (f *Docx) addImageRelation(m Media) string {
	rel := Relationship{
		ID:     "rId" + strconv.Itoa(int(atomic.AddUintptr(&f.rID, 1))),
		Type:   REL_IMAGE,
		Target: "media/" + m.Name,
	}

	f.docRelation.Relationship = append(f.docRelation.Relationship, rel)

	return rel.ID
}

// ReferTarget gets the target for a reference
func (f *Docx) ReferTarget(id string) (string, error) {
	for _, a := range f.docRelation.Relationship {
		if a.ID == id {
			return a.Target, nil
		}
	}
	return "", ErrRefIDNotFound
}

// ReferID gets the rId from target
func (f *Docx) ReferID(target string) (string, error) {
	for _, a := range f.docRelation.Relationship {
		if a.Target == target {
			return a.ID, nil
		}
	}
	return "", ErrRefIDNotFound
}

func (f *Docx) findRelationshipByID(id string) *Relationship {
	for i := range f.docRelation.Relationship {
		if f.docRelation.Relationship[i].ID == id {
			return &f.docRelation.Relationship[i]
		}
	}
	return nil
}

func (f *Docx) findRelationshipByTypeTarget(typ, target string) *Relationship {
	norm := normalizeRelTarget(target)
	for i := range f.docRelation.Relationship {
		r := &f.docRelation.Relationship[i]
		if r.Type == typ && normalizeRelTarget(r.Target) == norm {
			return r
		}
	}
	return nil
}

func normalizeRelTarget(target string) string {
	target = strings.TrimSpace(strings.TrimPrefix(target, "./"))
	target = strings.TrimPrefix(target, "/")
	target = strings.TrimPrefix(target, "word/")
	return target
}

func (f *Docx) ensureInternalPartRelation(relType, target, currentRID string) string {
	target = normalizeRelTarget(target)
	if currentRID != "" {
		if rel := f.findRelationshipByID(currentRID); rel != nil {
			rel.Type = relType
			rel.Target = target
			rel.TargetMode = ""
			return rel.ID
		}
	}
	if rel := f.findRelationshipByTypeTarget(relType, target); rel != nil {
		return rel.ID
	}
	rel := Relationship{
		ID:     "rId" + strconv.Itoa(int(atomic.AddUintptr(&f.rID, 1))),
		Type:   relType,
		Target: target,
	}
	f.docRelation.Relationship = append(f.docRelation.Relationship, rel)
	return rel.ID
}
