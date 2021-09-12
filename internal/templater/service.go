package templater

import (
	"context"
	"fmt"
	"os"

	"github.com/go-kit/kit/log"
	"github.com/go-kit/kit/log/level"
	"github.com/tealeg/xlsx/v3"
	"github.com/xuri/excelize/v2"
)

type fillInFunc func(ctx context.Context, template, tmpFile string, payload interface{}) error

type path interface {
	Template(name string) string
	TmpFile(name string) string
}

type parser interface {
	Type(filename string) (string, error)
}

type placeholder interface {
	Is(str string) bool
	GetValue(payload interface{}, placeholder string) (t string, value interface{}, err error)
}

type qrcode interface {
	Create(str string, size int) ([]byte, error)
}

type handlerPlaceholderHandler func(cell *xlsx.Cell, value interface{}) (err error)

type service struct {
	fillIn     map[string]fillInFunc
	keyHandler map[string]handlerPlaceholderHandler

	path        path
	parser      parser
	placeholder placeholder
	qrcode      qrcode

	logger log.Logger
}

func NewService(
	path path,
	parser parser,
	placeholder placeholder,
	qrcode qrcode,

	logger log.Logger,
) Service {
	s := &service{
		path:        path,
		parser:      parser,
		placeholder: placeholder,
		qrcode:      qrcode,
		logger:      logger,
	}

	s.fillIn = map[string]fillInFunc{
		"xlsx": s.Xlsx,
	}

	s.keyHandler = map[string]handlerPlaceholderHandler{
		FieldNameType: s.fieldNameKyeHandler,
		ArrayType:     s.arrayKeyHandler,
	}
	return s
}

// FillIn fills template by req.
func (s *service) FillIn(ctx context.Context, req Request) (res Response, err error) {
	logger := log.WithPrefix(s.logger, "method", "FillIn", "uuid", req.UUID)

	templateType, err := s.parser.Type(req.Template)
	if err != nil {
		level.Error(logger).Log("msg", "template type", "err", err)
		return
	}
	fillIn, isExist := s.fillIn[templateType]
	if !isExist {
		level.Error(logger).Log("msg", "unknown type", "type", templateType)
		err = errUnknownTemplateType
		return
	}

	res = Response{
		UUID:   req.UUID,
		UserID: req.UserID,
	}

	templatePath, tmpFilePath := s.path.Template(req.Template), s.path.TmpFile(req.Template)
	if err = fillIn(ctx, templatePath, tmpFilePath, req.Payload); err != nil {
		level.Error(logger).Log("msg", "fill in template", "template", templateType, "err", err)
		return
	}

	defer os.Remove(tmpFilePath)
	if res.Document, err = os.ReadFile(tmpFilePath); err != nil {
		level.Error(logger).Log("msg", "read tmp file", "tmp file name", tmpFilePath, "err", err)
	}
	return
}

type qrcodeInfo struct {
	data   [][]byte
	sheet  string
	rowNum int
	colNum int
	row    *xlsx.Row
}

func (s *service) Xlsx(ctx context.Context, template, tmpFile string, payload interface{}) error {
	file, err := xlsx.OpenFile(template)
	if err != nil {
		return err
	}

	if len(file.Sheets) == 0 {
		err = fmt.Errorf("template file has wrong number of sheets: %d", len(file.Sheets))
		return err
	}

	var (
		qrcodeList []qrcodeInfo
		sheet      = file.Sheets[0]
	)
	for rowIdx := 0; rowIdx < sheet.MaxRow; rowIdx++ {
		for colIdx := 0; colIdx < sheet.MaxCol; colIdx++ {
			row, _ := sheet.Row(rowIdx)
			cell := row.GetCell(colIdx)
			if s.placeholder.Is(cell.Value) {
				placeholder := cell.Value
				placeholderType, value, err := s.placeholder.GetValue(payload, placeholder)
				if err != nil {
					return err
				}

				if keyHandler, ok := s.keyHandler[placeholderType]; ok {
					if err = keyHandler(cell, value); err != nil {
						return fmt.Errorf("placeholder: %s err: %s", placeholder, err)
					}
				}

				// TODO:
				// add to xlsx lib inserting picture
				if placeholderType == QRCodeType {
					qrcodeArr, ok := value.([]interface{})
					if !ok {
						err = fmt.Errorf("wrong type in placeholder: %s", placeholder)
						return err
					}
					q := qrcodeInfo{
						sheet:  sheet.Name,
						rowNum: rowIdx + 1,
						colNum: colIdx,
						row:    row,
					}
					for _, qrcodeStr := range qrcodeArr {
						var (
							data []byte
							size = int(row.GetHeight() * 1.33)
						)
						str, ok := qrcodeStr.(string)
						if !ok {
							return fmt.Errorf("wrong type in placeholder: %s", placeholder)
						}
						if data, err = s.qrcode.Create(str, size); err != nil {
							return fmt.Errorf("qrcode generate: %s", err)
						}
						q.data = append(q.data, data)
					}
					qrcodeList = append(qrcodeList, q)
				}
			}
		}
	}

	if err = file.Save(tmpFile); err != nil {
		return fmt.Errorf("save template to tmp file err: %s", err)
	}

	if err = s.insertQRcode(tmpFile, qrcodeList); err != nil {
		return fmt.Errorf("insert qr codes err: %s", err)
	}
	return err
}

func (s *service) fieldNameKyeHandler(cell *xlsx.Cell, value interface{}) (err error) {
	cell.SetValue(value)
	return
}

func (s *service) arrayKeyHandler(cell *xlsx.Cell, value interface{}) error {
	sheet := cell.Row.Sheet
	colIdx, rowNum := cell.GetCoordinates()
	pIdx := rowNum + 1
	pRow, _ := sheet.Row(pIdx)

	array, ok := value.([]interface{})
	if !ok {
		return fmt.Errorf("wrong type array")
	}
	for i, item := range array {
		cRow, _ := sheet.AddRowAtIndex(pIdx + 2 + i)
		if pRow.GetHeight() != 0 {
			cRow.SetHeight(pRow.GetHeight())
		}
		for j := colIdx; j < sheet.MaxCol; j++ {
			pCell := pRow.GetCell(j)
			cCell := cRow.GetCell(j)
			cCell.SetStyle(pCell.GetStyle())
			cCell.Merge(pCell.HMerge, pCell.VMerge)
			placeholderType, value, err := s.placeholder.GetValue(item, pCell.Value)
			if err != nil {
				return err
			}
			if placeholderType == FieldNameType {
				cCell.SetValue(value)
			}
		}
	}
	if len(array) == 0 {
		sheet.RemoveRowAtIndex(colIdx - 1)
	}
	sheet.RemoveRowAtIndex(colIdx)
	sheet.RemoveRowAtIndex(colIdx)
	return nil
}

func (s *service) insertQRcode(file string, qrcodeList []qrcodeInfo) (err error) {
	f, _ := excelize.OpenFile(file)
	for _, qrcode := range qrcodeList {
		colNum := qrcode.colNum
		for _, qrData := range qrcode.data {
			cellName := fmt.Sprintf("%c%d", 'A'+colNum, qrcode.rowNum)
			if err = f.AddPictureFromBytes(qrcode.sheet, cellName, "", "", ".png", qrData); err != nil {
				return
			}
			f.SetCellValue(qrcode.sheet, cellName, "")
			colNum += qrcode.row.GetCell(colNum).HMerge + 1
		}
	}
	f.Save()
	return
}
