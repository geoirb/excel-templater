package templater

import (
	"errors"
)

var (
	errUnknownTemplateType = errors.New("unknown template type")
)

// if placeholderType == FieldNameType {
// 	cell.SetValue(value)
// }

// if placeholderType == ArrayType {
// 	pRow, _ := sheet.Row(rowIdx + 1)
// 	array, ok := value.([]interface{})
// 	if !ok {
// 		err = fmt.Errorf("wrong type in placeholder: %s", placeholder)
// 		return err
// 	}
// 	for i, item := range array {
// 		cRow, _ := sheet.AddRowAtIndex(rowIdx + 2 + i)
// 		if pRow.GetHeight() != 0 {
// 			cRow.SetHeight(pRow.GetHeight())
// 		}
// 		for j := colIdx; j < sheet.MaxCol; j++ {
// 			pCell := pRow.GetCell(j)
// 			cCell := cRow.GetCell(j)
// 			cCell.SetStyle(pCell.GetStyle())
// 			cCell.Merge(pCell.HMerge, pCell.VMerge)
// 			if placeholderType, value, err = s.placeholder.GetValue(item, pCell.Value); err != nil {
// 				return err
// 			}
// 			if placeholderType == FieldNameType {
// 				cCell.SetValue(value)
// 			}
// 		}
// 	}
// 	if len(array) == 0 {
// 		sheet.RemoveRowAtIndex(rowIdx - 1)
// 		rowIdx--
// 	}
// 	sheet.RemoveRowAtIndex(rowIdx)
// 	sheet.RemoveRowAtIndex(rowIdx)
// 	rowIdx += len(array) - 2
// }
