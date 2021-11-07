# EXCEL-TEMPLATER

## Description

The simple templater for filling in excel (.xlsx .xls) file with data.

Templater supports:
* Insert simple values
* Insert tabels
* Insert qr codes
* Insert list of qr codes
* Insert png images.
* Defaults values (turn on by flag)
* Fill in multipage file

Data for filling in must be in serializing format

## Placeholders

For correct input data in excel file, file must have placeholders. The placeholder shows the template engine where and what data will be inserted and sets the style of inserted data.

Placeholder is string between `{ }`, string consisting of keys`([_a-zA-Z0-9]+)`, keys are separated by `:`.
The placeholder shows the template engine:
- style of inserted data - stile of placeholder
- value of inserted data - set of keys in placeholder must be a path to value in data for filling in
- type of inserted data - last key

### Supported types of inserted data 

| type          | key         | default value | type of value in data for filling | describe                                                                                                                                                     |
|---------------|-------------|---------------|-----------------------------------|--------------------------------------------------------------------------------------------------------------------------------------------------------------|
| simple        |             | " "           | any type                          | the value will be inserted along the path from the placeholder                                                                                               |
| qr code       | qr_code     | " "           | not empty string                  | the qr code will be generated from the value, and it will be inserted in a cell, height of qr code will be the equal height of the cell with the placeholder |
| qr code array | qr_code_row |               | array of not empty strings        | the array of qr codes will be inserted in a row, starting with the placeholder cell                                                                          |
| image         | image       | transparent pixel    			| base64-encoded PNG |   
| table         | table       |                      			| array of objects   | __!!! row with this placeholder will be deleted!!!__ Each object from the array will be converted to the generated table row. For right converting in next row after row with table placeholder's cell must have placeholders row, this placeholders needs for inserting values in columns of the generated table, templater find value in each object of the array|

### Examples 

Simple:

_Data for filling_

```json
{
  "field_1": "value_for_A1",
  "field_2": "value_for_A2",
  "group_1": {
    "field_1": "value_for_D2",
    "subgroup": {
      "field": "value_for_C4"
    }
}
```

_Template_

![simple_value_template](images/simple_value_template.png)

_Result_

![simple_value_result](images/simple_value_result.png)

Qr code:

_Data for filling_

```json
{
  	"urls": [
		"https://t.me/geoirb",
		"https://github.com/geoirb",
		"https://github.com/geoirb/excel-templater/blob/master/README.md"
	],
	"url": "https://t.me/geoirb"
}
```

_Template_

![qr_code_template](images/qr_code_template.png)

_Result_

![qr_code_result](images/qr_code_result.png)


Table:

_Data for filling_

```json
{
	"group_1": {
		"table_example": [
			{
				"column_1": 1.1,
				"column_2": "first_row",
				"column_3": "first_row",
				"column_4": 0.0
			},
			{
				"column_1": 2.2,
				"column_2": "second_row",
				"column_3": "second_row",
				"column_4": 0.1
			},
			{
				"column_1": 3.3,
				"column_2": "third_row",
				"column_3": "third_row",
				"column_4": 0.2
			}
		]
	}
}
```

_Template_

![table_template](images/table_template.png)

_Result_

![table_result](images/table_result.png)


## Get start

```golang

import (
...

	"github.com/geoirb/excel-templater"
...
)

func main() {
	// flag useDefault turn on default values
	templater := excel.NewTemplater(useDefault)

	var payload interface{}
	// deserializing inserted data 
	if err := json.Unmarshal(data, &payload); err != nil {
		panic(err)
	}

	// templateFile - path to template
	r, err := templater.FillIn(templateFile, payload)
	if err != nil {
		panic(err)
	}
}
```

## Gratitude

- github.com/qax-os/excelize
- github.com/skip2/go-qrcode