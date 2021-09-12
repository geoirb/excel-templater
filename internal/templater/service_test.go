package templater_test

import (
	"context"
	"encoding/json"
	"os"
	"testing"

	"github.com/stretchr/testify/assert"

	"github.com/geoirb/go-templater/internal/placeholder"
	"github.com/geoirb/go-templater/internal/qrcode"
	"github.com/geoirb/go-templater/internal/templater"
)

func TestXLSX(t *testing.T) {
	placeholder, err := placeholder.New()
	assert.NoError(t, err)

	qrcode := qrcode.NewCreator()
	svc := templater.NewService(
		nil,
		nil,
		placeholder,
		qrcode,
		nil,
	)

	data, err := os.ReadFile("/home/geoirb/project/go/geoirb/templater/_path_to_template/payload.json")
	assert.NoError(t, err)

	var payload interface{}
	json.Unmarshal(data, &payload)

	err = svc.Xlsx(
		context.Background(),
		"/home/geoirb/project/go/geoirb/templater/_path_to_template/template.xlsx",
		"/home/geoirb/project/go/geoirb/templater/_path_to_template/result.xlsx",
		payload,
	)
	assert.NoError(t, err)
}
