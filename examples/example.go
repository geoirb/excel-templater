package main

import (
	"context"
	"encoding/json"
	"io/ioutil"

	"githab.com/geoirb/xlsx-templater"
	"github.com/xuri/excelize/v2"
)

func main() {
	templater := xlsx.NewTemplater(true)

	data, err := ioutil.ReadFile("examples/payload.json")
	if err != nil {
		panic(err)
	}
	var payload interface{}
	json.Unmarshal(data, &payload)

	r, err := templater.FillIn(context.Background(), "examples/template.xlsx", payload)
	if err != nil {
		panic(err)
	}
	file, err := excelize.OpenReader(r)
	if err != nil {
		panic(err)
	}

	if err = file.SaveAs("examples/result.xlsx"); err != nil {
		panic(err)
	}
}
