package xlsx_test

import (
	"context"
	"encoding/json"
	"os"
	"testing"

	"github.com/stretchr/testify/assert"

	"github.com/geoirb/go-templater/internal/placeholder"
	"github.com/geoirb/go-templater/internal/qrcode"
	"github.com/geoirb/go-templater/internal/xlsx"
)

func TestFillIn(t *testing.T) {
	placeholder, err := placeholder.New(true)
	assert.NoError(t, err)

	qrcode := qrcode.NewCreator()
	svc := xlsx.NewFacade(
		placeholder,
		qrcode,
	)

	data, err := os.ReadFile("/home/geoirb/project/go/geoirb/templater/.vscode/payload.json")
	assert.NoError(t, err)

	var payload interface{}
	err = json.Unmarshal(data, &payload)
	assert.NoError(t, err)
	_, err = svc.FillIn(
		context.Background(),
		"/home/geoirb/project/go/geoirb/templater/.vscode/template.xlsx",
		payload,
	)
	assert.NoError(t, err)
}
