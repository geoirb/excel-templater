package xlsx

import (
	"fmt"

	"github.com/geoirb/go-templater/internal/excelize"
)

func (s *Facade) fieldNameKyeHandler(file *excelize.File, sheetIdx int, rowIdx *int, colIdx int, value interface{}) (err error) {
	sheet := file.GetSheetName(sheetIdx)
	axis, _ := excelize.CoordinatesToCellName(colIdx+1, *rowIdx+1)
	return file.SetCellValue(sheet, axis, value)
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
		axis, _ := excelize.CoordinatesToCellName(colIdx+1, *rowIdx+1)
		if err = file.AddPictureFromBytes(sheet, axis, "", "", ".png", data); err != nil {
			return
		}
		file.SetCellValue(sheet, axis, "")
		colNum, _, _ := file.GetNumMergeCell(sheet, axis)
		colIdx += colNum
	}
	return
}

func (s *Facade) imageHandler(file *excelize.File, sheetIdx int, rowIdx *int, colIdx int, value interface{}) (err error) {
	sheet := file.GetSheetName(sheetIdx)
	axis, _ := excelize.CoordinatesToCellName(colIdx+1, *rowIdx+1)
	data, ok := value.(string)
	if !ok {
		return fmt.Errorf("imageHandler: wrong type payload")
	}
	return file.AddPictureFromBytes(sheet, axis, "", "", ".png", []byte(data))
}
