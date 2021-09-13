package placeholder

import (
	"fmt"
	"regexp"

	"github.com/geoirb/go-templater/internal/xlsx"
)

// Placeholder .,,
type Placeholder struct {
	placeholderGroupReg *regexp.Regexp
	placeholderReg      *regexp.Regexp
	arrayRegexp         *regexp.Regexp
	qrCodeRegexp        *regexp.Regexp
	imageReqexp         *regexp.Regexp

	valuesAreRequired bool
}

// New ...
func New(
	valuesAreRequired bool,
) (f *Placeholder, err error) {
	f = &Placeholder{
		valuesAreRequired: valuesAreRequired,
	}
	if f.placeholderGroupReg, err = regexp.Compile(placeholderGroupRegexp); err != nil {
		return
	}
	if f.placeholderReg, err = regexp.Compile(placeholderReqexp); err != nil {
		return
	}
	if f.arrayRegexp, err = regexp.Compile(arrayRegexp); err != nil {
		return
	}

	if f.qrCodeRegexp, err = regexp.Compile(qrCodeRegexp); err != nil {
		return
	}
	f.imageReqexp, err = regexp.Compile(imageReqexp)
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
		placeholderType = xlsx.FieldNameType
		if p.arrayRegexp.Match([]byte(key)) {
			placeholderType = xlsx.ArrayType
			return
		}
		if p.qrCodeRegexp.Match([]byte(key)) {
			placeholderType = xlsx.QRCodeType
		}
		if p.imageReqexp.Match([]byte(key)) {
			placeholderType = xlsx.ImageType
		}
		value, ok = p.value(value, key)

		if !ok {
			err = fmt.Errorf("wrong format payload: not found key %s in placeholder %s", key, placeholder)
			return
		}
	}
	return
}

func (p *Placeholder) value(payload interface{}, key string) (interface{}, bool) {
	if m, ok := payload.(map[string]interface{}); ok {
		value, isExist := m[key]
		if !isExist && !p.valuesAreRequired {
			value = ""
			isExist = true
		}
		return value, isExist
	}
	return nil, false
}
