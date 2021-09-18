package xlsx

import (
	"encoding/base64"
	"fmt"
	"strings"

	"github.com/xuri/excelize/v2"
)

func (s *Templater) fieldNameKyeHandler(file *excelize.File, sheetIdx int, rowIdx *int, colIdx int, value interface{}) (err error) {
	sheet := file.GetSheetName(sheetIdx)
	axis, _ := excelize.CoordinatesToCellName(colIdx+1, *rowIdx+1)
	return file.SetCellValue(sheet, axis, value)
}

func (s *Templater) arrayKeyHandler(file *excelize.File, sheetIdx int, rowIdx *int, _ int, value interface{}) error {
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

func (s *Templater) qrCodeHandler(file *excelize.File, sheetIdx int, rowIdx *int, colIdx int, value interface{}) (err error) {
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
		if data, err = s.qrcodeEncode(str, int(qrcodeSize)); err != nil {
			return fmt.Errorf("qrcode generate: %s", err)
		}
		axis, _ := excelize.CoordinatesToCellName(colIdx+1, *rowIdx+1)
		if err = file.AddPictureFromBytes(sheet, axis, "", "", ".png", data); err != nil {
			return
		}
		file.SetCellValue(sheet, axis, "")
		colNum, _, _ := getNumMergeCell(file, sheet, axis)
		colIdx += colNum
	}
	return
}

func (s *Templater) imageHandler(file *excelize.File, sheetIdx int, rowIdx *int, colIdx int, value interface{}) (err error) {
	sheet := file.GetSheetName(sheetIdx)
	axis, _ := excelize.CoordinatesToCellName(colIdx+1, *rowIdx+1)
	file.SetCellValue(sheet, axis, "")

	image, ok := value.(string)
	if !ok {
		return fmt.Errorf("imageHandler: wrong type payload")
	}
	i := strings.Index(image, ",")
	image = image[i+1:]
	if len(image) == 0 {
		image = defaultImage
	}
	imageBytes, _ := base64.StdEncoding.DecodeString(image)
	return file.AddPictureFromBytes(sheet, axis, "", "", ".png", imageBytes)
}

// For quick work add to github.com/xuri/excelize/v2 function:
// GetNumMergeCell provides a function to get the number of merged rows and columns by axis cell
// from a worksheet currently.
// func (f *File) GetNumMergeCell(sheet string, axis string) (colNum int, rowNum int, err error) {
// 	ws, err := f.workSheetReader(sheet)
// 	if err != nil {
// 		return
// 	}

// 	if ws.MergeCells != nil {
// 		for i := range ws.MergeCells.Cells {
// 			ref := ws.MergeCells.Cells[i].Ref
// 			cells := strings.Split(ref, ":")
// 			if cells[0] == axis {
// 				col1, row1, _ := CellNameToCoordinates(cells[0])
// 				col2, row2, _ := CellNameToCoordinates(cells[1])
// 				colNum, rowNum = col2-col1+1, row2-row1+1
// 				return
// 			}
// 		}
// 	}
// 	colNum = 1
// 	rowNum = 1
// 	return
// }
func getNumMergeCell(file *excelize.File, sheet string, axis string) (colNum int, rowNum int, err error) {
	mergedCells, err := file.GetMergeCells(sheet)
	if err != nil {
		return
	}
	for _, mergetCell := range mergedCells {
		if mergetCell.GetStartAxis() == axis {
			col1, row1, _ := excelize.CellNameToCoordinates(mergetCell.GetStartAxis())
			col2, row2, _ := excelize.CellNameToCoordinates(mergetCell.GetStartAxis())
			colNum, rowNum = col2-col1+1, row2-row1+1
			return
		}
	}
	colNum = 1
	rowNum = 1
	return
}
