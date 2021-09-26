package main

import (
	_ "embed"
	"encoding/json"

	"github.com/geoirb/excel-templater"
	"github.com/xuri/excelize/v2"
)

var (
	templateFile      = "/home/geoirb/project/go/geoirb/excel-templater/example/template.xlsx"
	resultFile        = "example/result.xlsx"
	valuesAreRequired = true

	//go:embed payload.json
	data []byte
)

func main() {
	templater := excel.NewTemplater(valuesAreRequired)

	var payload interface{}
	if err := json.Unmarshal(data, &payload); err != nil {
		panic(err)
	}

	r, err := templater.FillIn(templateFile, payload)
	if err != nil {
		panic(err)
	}
	file, err := excelize.OpenReader(r)
	if err != nil {
		panic(err)
	}

	if err = file.SaveAs(resultFile); err != nil {
		panic(err)
	}
}
