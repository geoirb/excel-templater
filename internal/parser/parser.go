package parser

import (
	"regexp"
)

// Parser ...
type Parser struct {
	typeRegexp *regexp.Regexp
}

// New ...
func New() (p *Parser, err error) {
	p = &Parser{}
	p.typeRegexp, err = regexp.Compile(typeRegexp)
	return
}

// Type returns type of file by filename.
func (p *Parser) Type(filename string) (string, error) {
	if submathList := p.typeRegexp.FindAllStringSubmatch(filename, -1); len(submathList) == 1 && len(submathList[0]) == 2 {
		return submathList[0][1], nil
	}
	return "", errTypeNotDefined
}
