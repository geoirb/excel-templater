package excel

import (
	"bytes"
	"fmt"
	"io"

	"github.com/xuri/excelize/v2"
)

const (
	// placeholder types.
	fieldNameType = "field_name"
	tableType     = "table"
	qrCodeType    = "qr_code"
	qrCodeRowType = "qr_code_row"
	imageType     = "image"
)

type qrcodeEncodeFunc func(str string, pixels int) ([]byte, error)

type placeholderHandler func(file *excelize.File, sheet string, rowNumb, colIdx *int, value interface{}) (err error)

// templater для заполнения excel файлов.
type templater struct {
	keyHandler map[string]placeholderHandler

	placeholder  *placeholder
	qrcodeEncode qrcodeEncodeFunc
}

// NewTemplater
// useDefault - flag for turn on default values.
func NewTemplater(
	useDefault bool,
) *templater {
	f := &templater{
		placeholder:  newPlaceholdParser(useDefault),
		qrcodeEncode: encode,
	}
	f.keyHandler = map[string]placeholderHandler{
		fieldNameType: f.fieldNameKyeHandler,
		tableType:     f.tableKeyHandler,
		qrCodeType:    f.qrCodeHandler,
		qrCodeRowType: f.qrCodeRowHandler,
		imageType:     f.imageHandler,
	}
	return f
}

// FillIn file by templatePath path with payload data.
func (t *templater) FillIn(templatePath string, payload interface{}) (r io.Reader, err error) {
	f, err := excelize.OpenFile(templatePath)
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

//  filling in on sheet.
func (t *templater) fillInSheet(file *excelize.File, sheet string, payload interface{}) (err error) {
	rows, err := file.GetRows(sheet)
	if err != nil {
		return
	}

	for rowIdx := 0; rowIdx < len(rows); rowIdx++ {
		for colIdx, cellValue := range rows[rowIdx] {

			var (
				placeholderType string
				value           interface{}
			)
			placeholderType, value, err = t.placeholder.GetValue(payload, cellValue)
			if err != nil {
				return
			}

			if keyHandler, ok := t.keyHandler[placeholderType]; ok {
				if err = keyHandler(file, sheet, &rowIdx, &colIdx, value); err != nil {
					err = fmt.Errorf("placeholder: %s err: %s", cellValue, err)
					return
				}
				if rows, err = file.GetRows(sheet); err != nil {
					err = fmt.Errorf("placeholder: %s err: %s", cellValue, err)
					return
				}
			}

		}
	}
	return
}
