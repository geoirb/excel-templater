package excel

import (
	"fmt"
	"regexp"
)

const (
	placeholderGroupRegexp = "{([_a-zA-Z0-9:]+)}$"
	placeholderReqexp      = "([_a-zA-Z0-9]+)"
	tableRegexp            = "table$"
	qrCodeRegexp           = "qr_code_[_a-zA-Z0-9]+"
	imageReqexp            = "image_[_a-zA-Z0-9]+"
)

type placeholder struct {
	placeholderGroupReg *regexp.Regexp
	placeholderReg      *regexp.Regexp
	tableRegexp         *regexp.Regexp
	qrCodeRegexp        *regexp.Regexp
	imageReqexp         *regexp.Regexp

	valuesAreRequired bool
}

func newPlaceholdParser(
	valuesAreRequired bool,
) *placeholder {
	return &placeholder{
		valuesAreRequired:   valuesAreRequired,
		placeholderGroupReg: regexp.MustCompile(placeholderGroupRegexp),
		placeholderReg:      regexp.MustCompile(placeholderReqexp),
		tableRegexp:         regexp.MustCompile(tableRegexp),
		qrCodeRegexp:        regexp.MustCompile(qrCodeRegexp),
		imageReqexp:         regexp.MustCompile(imageReqexp),
	}
}

// Is returns true if str is playsholder.
func (p *placeholder) Is(str string) bool {
	return p.placeholderGroupReg.Match([]byte(str))
}

// GetValue from data by placeholder.
func (p *placeholder) GetValue(payload interface{}, placeholder string) (placeholderType string, value interface{}, err error) {
	value = payload
	keys := p.placeholderReg.FindAllString(placeholder, -1)
	for _, key := range keys {
		var ok bool
		switch key {
		case tableType, qrCodeType, imageType:
			placeholderType = key
			return
		default:
			placeholderType = fieldNameType
			if value, ok = p.value(value, key); !ok {
				err = fmt.Errorf("wrong payload: not found key %s from placeholder %s", key, placeholder)
				return
			}
		}
	}
	return
}

func (p *placeholder) value(payload interface{}, key string) (interface{}, bool) {
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
