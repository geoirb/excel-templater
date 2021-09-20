package xlsx

import (
	"bytes"
	"context"
	"fmt"
	"io"

	"github.com/xuri/excelize/v2"
)

const (

	// Placeholder types.
	FieldNameType = "field_name"
	ArrayType     = "array"
	QRCodeType    = "qr_code"
	ImageType     = "image"
)

type placeholderParser interface {
	Is(str string) bool
	GetValue(payload interface{}, placeholder string) (t string, value interface{}, err error)
}

type qrcodeEncodeFunc func(str string, pixels int) ([]byte, error)

type placeholderHandler func(file *excelize.File, sheet string, rowNumb, colIdx *int, value interface{}) (err error)

// Templater for xlsx.
type Templater struct {
	keyHandler map[string]placeholderHandler

	placeholder  placeholderParser
	qrcodeEncode qrcodeEncodeFunc
}

// NewTemplater for xlsx.
func NewTemplater(
	valuesAreRequired bool,
) *Templater {
	f := &Templater{
		placeholder:  newPlaceholdParser(valuesAreRequired),
		qrcodeEncode: encode,
	}
	f.keyHandler = map[string]placeholderHandler{
		FieldNameType: f.fieldNameKyeHandler,
		ArrayType:     f.arrayKeyHandler,
		QRCodeType:    f.qrCodeHandler,
		ImageType:     f.imageHandler,
	}
	return f
}

// FillIn template from payload.
func (t *Templater) FillIn(ctx context.Context, template string, payload interface{}) (r io.Reader, err error) {
	f, err := excelize.OpenFile(template)
	if err != nil {
		return
	}

	sheets := f.GetSheetList()
	for _, sheet := range sheets {
		if err = t.fillInSheet(f, sheet, payload); err != nil {
			err = fmt.Errorf("sheet: %s %s", sheet, err)
			return
		}
	}

	var result bytes.Buffer
	if err = f.Write(&result); err != nil {
		err = fmt.Errorf("save template to tmp file err: %s", err)
		return
	}
	r = &result
	return
}

func (t *Templater) fillInSheet(f *excelize.File, sheet string, payload interface{}) (err error) {
	rows, err := f.GetRows(sheet)
	if err != nil {
		return
	}

	for rowIdx := 0; rowIdx < len(rows); rowIdx++ {
		for colIdx, cellValue := range rows[rowIdx] {
			if t.placeholder.Is(cellValue) {
				var (
					placeholderType string
					value           interface{}
				)
				placeholderType, value, err = t.placeholder.GetValue(payload, cellValue)
				if err != nil {
					return
				}

				if keyHandler, ok := t.keyHandler[placeholderType]; ok {
					if err = keyHandler(f, sheet, &rowIdx, &colIdx, value); err != nil {
						err = fmt.Errorf("placeholder: %s err: %s", cellValue, err)
						return
					}
					if rows, err = f.GetRows(sheet); err != nil {
						err = fmt.Errorf("placeholder: %s err: %s", cellValue, err)
						return
					}
				}
			}
		}
	}
	return
}
