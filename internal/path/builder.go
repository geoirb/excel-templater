package path

import (
	"fmt"
	"os"
	"strings"
)

// Builder path.
type Builder struct {
	templateDir string
	tmpDir      string
	uuidFunc    func() string
}

// NewBuilder ...
func NewBuilder(
	templateDir string,
	tmpDir string,
	uuidFunc func() string,
) (*Builder, error) {
	if _, err := os.Stat(templateDir); os.IsNotExist(err) {
		return nil, fmt.Errorf("path %s is not exist", templateDir)
	}
	if !strings.HasSuffix(templateDir, "/") {
		templateDir += "/"
	}
	
	if _, err := os.Stat(tmpDir); os.IsNotExist(err) {
		return nil, fmt.Errorf("path %s is not exist", templateDir)
	}
	if !strings.HasSuffix(tmpDir, "/") {
		tmpDir += "/"
	}

	return &Builder{
		templateDir: templateDir,
		tmpDir:      tmpDir,
		uuidFunc:    uuidFunc,
	}, nil
}

// Template returns path to template by name.
func (b *Builder) Template(name string) string {
	return b.templateDir + name
}

// TmpFile returns path to tmp file by name.
func (b *Builder) TmpFile(name string) string {
	return b.tmpDir + b.uuidFunc() + name
}
