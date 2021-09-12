package placeholder

import (
	"fmt"
	"regexp"

	"github.com/geoirb/go-templater/internal/templater"
)

// Placeholder .,,
type Placeholder struct {
	placeholderGroupReg *regexp.Regexp
	placeholderReg      *regexp.Regexp
	arrayRegexp         *regexp.Regexp
	qrCodeRegexp        *regexp.Regexp
}

// New ...
func New() (f *Placeholder, err error) {
	f = &Placeholder{}
	if f.placeholderGroupReg, err = regexp.Compile(placeholderGroupRegexp); err != nil {
		return
	}
	if f.placeholderReg, err = regexp.Compile(placeholderReqexp); err != nil {
		return
	}
	if f.arrayRegexp, err = regexp.Compile(arrayRegexp); err != nil {
		return
	}

	f.qrCodeRegexp, err = regexp.Compile(qrCodeRegexp)
	return
}

// Is returns true if str is playsholder.
func (p *Placeholder) Is(str string) bool {
	return p.placeholderGroupReg.Match([]byte(str))
}

// GetValue from data by placeholder.
func (p *Placeholder) GetValue(payload interface{}, placeholder string) (placeholderType string, value interface{}, err error) {
	value = payload
	keys := p.placeholderReg.FindAllString(placeholder, -1)
	for _, key := range keys {
		var ok bool
		placeholderType = templater.FieldNameType
		if p.arrayRegexp.Match([]byte(key)) {
			placeholderType = templater.ArrayType
			return
		}
		if p.qrCodeRegexp.Match([]byte(key)) {
			placeholderType = templater.QRCodeType
		}
		value, ok = p.value(value, key)

		if !ok {
			err = fmt.Errorf("wrong format payload: not found key %s in placeholder %s", key, placeholder)
			return
		}
	}
	return
}

func (*Placeholder) value(payload interface{}, key string) (interface{}, bool) {
	if m, ok := payload.(map[string]interface{}); ok {
		value, isExist := m[key]
		return value, isExist
	}
	return nil, false
}
