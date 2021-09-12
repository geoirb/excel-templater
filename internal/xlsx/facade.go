package xlsx

import (
	"bytes"
	"context"
	"fmt"
	"io"

	"github.com/geoirb/go-templater/internal/excelize"
)

type placeholder interface {
	Is(str string) bool
	GetValue(payload interface{}, placeholder string) (t string, value interface{}, err error)
}

type qrcode interface {
	Create(str string, size int) ([]byte, error)
}

type handlerPlaceholderHandler func(file *excelize.File, sheetIdx int, rowNumb *int, colIdx int, value interface{}) (err error)

// Facade for xlsx
type Facade struct {
	keyHandler map[string]handlerPlaceholderHandler

	placeholder placeholder
	qrcode      qrcode
}

func NewFacade(
	placeholder placeholder,
	qrcode qrcode,
) *Facade {
	f := &Facade{
		placeholder: placeholder,
		qrcode:      qrcode,
	}
	f.keyHandler = map[string]handlerPlaceholderHandler{
		FieldNameType: f.fieldNameKyeHandler,
		ArrayType:     f.arrayKeyHandler,
		QRCodeType:    f.qrCodeHandler,
	}
	return f
}

func (s *Facade) FillIn(ctx context.Context, template string, payload interface{}) (r io.Reader, err error) {
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
					continue
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

func (s *Facade) fieldNameKyeHandler(file *excelize.File, sheetIdx int, rowIdx *int, colIdx int, value interface{}) (err error) {
	sheet := file.GetSheetName(sheetIdx)
	axis := fmt.Sprintf("%s%d", getColByIdx(colIdx), *rowIdx+1)
	file.SetCellValue(sheet, axis, value)
	return
}

func (s *Facade) arrayKeyHandler(file *excelize.File, sheetIdx int, rowIdx *int, _ int, value interface{}) error {
	rowNumb := *rowIdx + 1
	sheet := file.GetSheetName(sheetIdx)
	rows, _ := file.GetRows(sheet)

	array, ok := value.([]interface{})
	if !ok {
		return fmt.Errorf("arrayKeyHandler: wrong type payload")
	}
	hRowNumb := rowNumb + 1
	hRow := rows[hRowNumb-1]
	for i, item := range array {
		file.DuplicateRowTo(sheet, hRowNumb, hRowNumb+i+1)
		for idx, cellValue := range hRow {
			placeholderType, value, err := s.placeholder.GetValue(item, cellValue)
			if err != nil {
				return err
			}
			if placeholderType == FieldNameType {
				rowIdx := hRowNumb + i
				s.fieldNameKyeHandler(file, sheetIdx, &rowIdx, idx, value)
			}
		}
	}

	file.RemoveRow(sheet, rowNumb)
	file.RemoveRow(sheet, rowNumb)
	if len(array) == 0 {
		file.RemoveRow(sheet, *rowIdx)
		*rowIdx--
	}
	*rowIdx = *rowIdx + len(array) - 2
	return nil
}

func (s *Facade) qrCodeHandler(file *excelize.File, sheetIdx int, rowIdx *int, colIdx int, value interface{}) (err error) {
	sheet := file.GetSheetName(sheetIdx)
	colSize, _ := file.GetRowHeight(sheet, *rowIdx+1)
	qrcodeSize := colSize * 1.333
	qrcodeArr, ok := value.([]interface{})
	if !ok {
		err = fmt.Errorf("qrCodeHandler: wrong type payload")
		return err
	}

	for _, qrcodeStr := range qrcodeArr {
		str, ok := qrcodeStr.(string)
		if !ok {
			return fmt.Errorf("qrCodeHandler: wrong type payload")
		}
		var data []byte
		if data, err = s.qrcode.Create(str, int(qrcodeSize)); err != nil {
			return fmt.Errorf("qrcode generate: %s", err)
		}
		axis := fmt.Sprintf("%s%d", getColByIdx(colIdx), *rowIdx+1)
		if err = file.AddPictureFromBytes(sheet, axis, "", "", ".png", data); err != nil {
			return
		}
		file.SetCellValue(sheet, axis, "")
		h, _ := file.GetHMergeCell(sheet, axis)
		colIdx += h
	}
	return
}
