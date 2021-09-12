package path

import (
	"fmt"
	"os"
	"strings"
)

// Builder path.
type Builder struct {
	templateDir string
}

// NewBuilder ...
func NewBuilder(
	templateDir string,
) (*Builder, error) {
	if _, err := os.Stat(templateDir); os.IsNotExist(err) {
		return nil, fmt.Errorf("path %s is not exist", templateDir)
	}
	if !strings.HasSuffix(templateDir, "/") {
		templateDir += "/"
	}

	return &Builder{
		templateDir: templateDir,
	}, nil
}

// Template returns path to template by name.
func (b *Builder) Template(name string) string {
	return b.templateDir + name
}
