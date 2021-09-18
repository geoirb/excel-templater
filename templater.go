package xlsx

import (
	"bytes"
	"context"
	"fmt"
	"io"

	"github.com/xuri/excelize/v2"
)

type placeholder interface {
	Is(str string) bool
	GetValue(payload interface{}, placeholder string) (t string, value interface{}, err error)
}

type qrcodeEncodeFunc func(str string, size int) ([]byte, error)

type handlerPlaceholderHandler func(file *excelize.File, sheetIdx int, rowNumb *int, colIdx int, value interface{}) (err error)

// Templater for xlsx.
type Templater struct {
	keyHandler map[string]handlerPlaceholderHandler

	placeholder  placeholder
	qrcodeEncode qrcodeEncodeFunc
}

// NewTemplater for xlsx.
func NewTemplater(
	valuesAreRequired bool,
) *Templater {
	f := &Templater{
		placeholder:  newPlaceholdParser(valuesAreRequired),
		qrcodeEncode: Encode,
	}
	f.keyHandler = map[string]handlerPlaceholderHandler{
		FieldNameType: f.fieldNameKyeHandler,
		ArrayType:     f.arrayKeyHandler,
		QRCodeType:    f.qrCodeHandler,
		ImageType:     f.imageHandler,
	}
	return f
}

// FillIn template from payload.
func (s *Templater) FillIn(ctx context.Context, template string, payload interface{}) (r io.Reader, err error) {
	f, err := excelize.OpenFile(template)
	if err != nil {
		return
	}

	sheets := f.GetSheetList()
	if len(sheets) == 0 {
		err = fmt.Errorf("template file has wrong number of sheets: %d", len(sheets))
		return
	}

	sheetIdx := 0
	sheet := sheets[sheetIdx]

	rows, err := f.GetRows(sheet)
	if err != nil {
		return
	}
	for rowIdx := 0; rowIdx < len(rows); rowIdx++ {
		for colIdx, cellValue := range rows[rowIdx] {
			if s.placeholder.Is(cellValue) {
				var (
					placeholderType string
					value           interface{}
				)
				placeholderType, value, err = s.placeholder.GetValue(payload, cellValue)
				if err != nil {
					return
				}

				if keyHandler, ok := s.keyHandler[placeholderType]; ok {
					if err = keyHandler(f, sheetIdx, &rowIdx, colIdx, value); err != nil {
						err = fmt.Errorf("placeholder: %s err: %s", cellValue, err)
						return
					}
					rows, err = f.GetRows(sheet)
					if err != nil {
						err = fmt.Errorf("placeholder: %s err: %s", cellValue, err)
						return
					}
				}
			}
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
