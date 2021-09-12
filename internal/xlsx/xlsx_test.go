package xlsx_test

import (
	"context"
	"encoding/json"
	"fmt"
	"io/ioutil"
	"os"
	"testing"

	"github.com/stretchr/testify/assert"

	"github.com/geoirb/go-templater/internal/placeholder"
	"github.com/geoirb/go-templater/internal/qrcode"
	"github.com/geoirb/go-templater/internal/xlsx"
)

func BenchmarkXLSX(b *testing.B) {
	placeholder, err := placeholder.New()
	assert.NoError(b, err)

	qrcode := qrcode.NewCreator()
	svc := xlsx.NewFacade(
		placeholder,
		qrcode,
	)

	data, err := os.ReadFile("/home/geoirb/project/go/geoirb/templater/_path_to_template/payload.json")
	assert.NoError(b, err)

	var payload interface{}
	json.Unmarshal(data, &payload)

	// for i := 0; i < b.N; i++ {
	r, _ := svc.FillIn(
		context.Background(),
		"/home/geoirb/project/go/geoirb/templater/_path_to_template/template.xlsx",
		payload,
	)
	os.Remove("/home/geoirb/project/go/geoirb/templater/_path_to_template/result.xlsx")
	fmt.Println(ioutil.ReadAll(r))
}
