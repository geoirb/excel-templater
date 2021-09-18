package xlsx

import (
	"fmt"
	"regexp"
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
func newPlaceholdParser(
	valuesAreRequired bool,
) *Placeholder {
	return &Placeholder{
		valuesAreRequired:   valuesAreRequired,
		placeholderGroupReg: regexp.MustCompile(placeholderGroupRegexp),
		placeholderReg:      regexp.MustCompile(placeholderReqexp),
		arrayRegexp:         regexp.MustCompile(arrayRegexp),
		qrCodeRegexp:        regexp.MustCompile(qrCodeRegexp),
		imageReqexp:         regexp.MustCompile(imageReqexp),
	}
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
		placeholderType = FieldNameType
		if p.arrayRegexp.Match([]byte(key)) {
			placeholderType = ArrayType
			return
		}
		if p.qrCodeRegexp.Match([]byte(key)) {
			placeholderType = QRCodeType
		}
		if p.imageReqexp.Match([]byte(key)) {
			placeholderType = ImageType
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
