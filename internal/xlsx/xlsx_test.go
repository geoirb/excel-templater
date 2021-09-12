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
	placeholder, err := placeholder.New(false)
	assert.NoError(t, err)

	qrcode := qrcode.NewCreator()
	svc := xlsx.NewFacade(
		placeholder,
		qrcode,
	)

	data, err := os.ReadFile("/home/geoirb/project/go/geoirb/templater/_path_to_template/payload.json")
	assert.NoError(t, err)

	var payload interface{}
	json.Unmarshal(data, &payload)

	_, err = svc.FillIn(
		context.Background(),
		"/home/geoirb/project/go/geoirb/templater/_path_to_template/template.xlsx",
		payload,
	)
	assert.NoError(t, err)
}
