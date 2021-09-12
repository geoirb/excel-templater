package placeholder

import (
	"encoding/json"
	"testing"

	"github.com/stretchr/testify/assert"

	"github.com/geoirb/go-templater/internal/xlsx"
)

const (
	testJson = `
	{
		"data_to_A1":"A1",
		"data_to_A2":"wrong",
		"data_to_B":{
		   "data_to_5":"B5",
		   "data_to_6":26.0
		},
		"data":[
		   [
				"wrong",
				3802
		   ],
		   [
				"wrong",
				3802,
				3802,
				3802,
				3802,
				"62"
		   ],
		   [
				"wrong",
				3802,
				"wrong",
				3802
		   ]
		],
		"qr_code_0": [
			"qr_code_01",
			"qr_code_02",
		]
	}
	`

	testPlaceholder = "{test_placeholder}"
	testStr         = "test-string"
)

var (
	nilInterface interface{}
	testCase     []struct {
		placeholder     string
		value           interface{}
		placeholderType string
	} = []struct {
		placeholder     string
		value           interface{}
		placeholderType string
	}{
		{
			placeholder:     "{data_to_A1}",
			value:           "A1",
			placeholderType: xlsx.FieldNameType,
		},
		{
			placeholder:     "{data_to_B:data_to_6}",
			value:           26.0,
			placeholderType: xlsx.FieldNameType,
		},
		{
			placeholder:     "{data:array}",
			placeholderType: xlsx.ArrayType,
		},
		{
			placeholder:     "{qr_code_0}",
			placeholderType: xlsx.QRCodeType,
		},
	}
)

func TestIs(t *testing.T) {
	p, err := New()
	assert.NoError(t, err)

	assert.True(t, p.Is(testPlaceholder))
	assert.False(t, p.Is(testStr))
}

func TestGetValue(t *testing.T) {
	var payload interface{}
	err := json.Unmarshal([]byte(testJson), &payload)
	assert.NoError(t, err)

	p, err := New()
	assert.NoError(t, err)

	for _, test := range testCase {
		actualType, actualValue, err := p.GetValue(payload, test.placeholder)
		assert.NoError(t, err)
		assert.Equal(t, test.placeholderType, actualType)
		if nilInterface != test.value {
			assert.Equal(t, test.value, actualValue)
		}
	}
}
