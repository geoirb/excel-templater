package excelize

import (
	"archive/zip"
	"bytes"
	"container/list"
	"encoding/xml"
	"fmt"
	"io"
	"regexp"
	"strconv"
	"strings"
	"unicode"
)

// ReadZipReader can be used to read the spreadsheet in memory without touching the
// filesystem.
func ReadZipReader(r *zip.Reader) (map[string][]byte, int, error) {
	var err error
	var docPart = map[string]string{
		"[content_types].xml":  "[Content_Types].xml",
		"xl/sharedstrings.xml": "xl/sharedStrings.xml",
	}
	fileList := make(map[string][]byte, len(r.File))
	worksheets := 0
	for _, v := range r.File {
		fileName := strings.Replace(v.Name, "\\", "/", -1)
		if partName, ok := docPart[strings.ToLower(fileName)]; ok {
			fileName = partName
		}
		if fileList[fileName], err = readFile(v); err != nil {
			return nil, 0, err
		}
		if strings.HasPrefix(fileName, "xl/worksheets/sheet") {
			worksheets++
		}
	}
	return fileList, worksheets, nil
}

// readXML provides a function to read XML content as string.
func (f *File) readXML(name string) []byte {
	if content, _ := f.Pkg.Load(name); content != nil {
		return content.([]byte)
	}
	if content, ok := f.streams[name]; ok {
		return content.rawData.buf.Bytes()
	}
	return []byte{}
}

// saveFileList provides a function to update given file content in file list
// of XLSX.
func (f *File) saveFileList(name string, content []byte) {
	f.Pkg.Store(name, append([]byte(XMLHeader), content...))
}

// Read file content as string in a archive file.
func readFile(file *zip.File) ([]byte, error) {
	rc, err := file.Open()
	if err != nil {
		return nil, err
	}
	dat := make([]byte, 0, file.FileInfo().Size())
	buff := bytes.NewBuffer(dat)
	_, _ = io.Copy(buff, rc)
	rc.Close()
	return buff.Bytes(), nil
}

// SplitCellName splits cell name to column name and row number.
//
// Example:
//
//     excelize.SplitCellName("AK74") // return "AK", 74, nil
//
func SplitCellName(cell string) (string, int, error) {
	alpha := func(r rune) bool {
		return ('A' <= r && r <= 'Z') || ('a' <= r && r <= 'z')
	}

	if strings.IndexFunc(cell, alpha) == 0 {
		i := strings.LastIndexFunc(cell, alpha)
		if i >= 0 && i < len(cell)-1 {
			col, rowstr := cell[:i+1], cell[i+1:]
			if row, err := strconv.Atoi(rowstr); err == nil && row > 0 {
				return col, row, nil
			}
		}
	}
	return "", -1, newInvalidCellNameError(cell)
}

// JoinCellName joins cell name from column name and row number.
func JoinCellName(col string, row int) (string, error) {
	normCol := strings.Map(func(rune rune) rune {
		switch {
		case 'A' <= rune && rune <= 'Z':
			return rune
		case 'a' <= rune && rune <= 'z':
			return rune - 32
		}
		return -1
	}, col)
	if len(col) == 0 || len(col) != len(normCol) {
		return "", newInvalidColumnNameError(col)
	}
	if row < 1 {
		return "", newInvalidRowNumberError(row)
	}
	return normCol + strconv.Itoa(row), nil
}

// ColumnNameToNumber provides a function to convert Excel sheet column name
// to int. Column name case insensitive. The function returns an error if
// column name incorrect.
//
// Example:
//
//     excelize.ColumnNameToNumber("AK") // returns 37, nil
//
func ColumnNameToNumber(name string) (int, error) {
	if len(name) == 0 {
		return -1, newInvalidColumnNameError(name)
	}
	col := 0
	multi := 1
	for i := len(name) - 1; i >= 0; i-- {
		r := name[i]
		if r >= 'A' && r <= 'Z' {
			col += int(r-'A'+1) * multi
		} else if r >= 'a' && r <= 'z' {
			col += int(r-'a'+1) * multi
		} else {
			return -1, newInvalidColumnNameError(name)
		}
		multi *= 26
	}
	if col > TotalColumns {
		return -1, ErrColumnNumber
	}
	return col, nil
}

// ColumnNumberToName provides a function to convert the integer to Excel
// sheet column title.
//
// Example:
//
//     excelize.ColumnNumberToName(37) // returns "AK", nil
//
func ColumnNumberToName(num int) (string, error) {
	if num < 1 {
		return "", fmt.Errorf("incorrect column number %d", num)
	}
	if num > TotalColumns {
		return "", ErrColumnNumber
	}
	var col string
	for num > 0 {
		col = string(rune((num-1)%26+65)) + col
		num = (num - 1) / 26
	}
	return col, nil
}

// CellNameToCoordinates converts alphanumeric cell name to [X, Y] coordinates
// or returns an error.
//
// Example:
//
//    excelize.CellNameToCoordinates("A1") // returns 1, 1, nil
//    excelize.CellNameToCoordinates("Z3") // returns 26, 3, nil
//
func CellNameToCoordinates(cell string) (int, int, error) {
	const msg = "cannot convert cell %q to coordinates: %v"

	colname, row, err := SplitCellName(cell)
	if err != nil {
		return -1, -1, fmt.Errorf(msg, cell, err)
	}
	if row > TotalRows {
		return -1, -1, fmt.Errorf("row number exceeds maximum limit")
	}
	col, err := ColumnNameToNumber(colname)
	return col, row, err
}

// CoordinatesToCellName converts [X, Y] coordinates to alpha-numeric cell
// name or returns an error.
//
// Example:
//
//    excelize.CoordinatesToCellName(1, 1) // returns "A1", nil
//    excelize.CoordinatesToCellName(1, 1, true) // returns "$A$1", nil
//
func CoordinatesToCellName(col, row int, abs ...bool) (string, error) {
	if col < 1 || row < 1 {
		return "", fmt.Errorf("invalid cell coordinates [%d, %d]", col, row)
	}
	sign := ""
	for _, a := range abs {
		if a {
			sign = "$"
		}
	}
	colname, err := ColumnNumberToName(col)
	return sign + colname + sign + strconv.Itoa(row), err
}

// boolPtr returns a pointer to a bool with the given value.
func boolPtr(b bool) *bool { return &b }

// intPtr returns a pointer to a int with the given value.
func intPtr(i int) *int { return &i }

// float64Ptr returns a pofloat64er to a float64 with the given value.
func float64Ptr(f float64) *float64 { return &f }

// stringPtr returns a pointer to a string with the given value.
func stringPtr(s string) *string { return &s }

// defaultTrue returns true if b is nil, or the pointed value.
func defaultTrue(b *bool) bool {
	if b == nil {
		return true
	}
	return *b
}

// parseFormatSet provides a method to convert format string to []byte and
// handle empty string.
func parseFormatSet(formatSet string) []byte {
	if formatSet != "" {
		return []byte(formatSet)
	}
	return []byte("{}")
}

// namespaceStrictToTransitional provides a method to convert Strict and
// Transitional namespaces.
func namespaceStrictToTransitional(content []byte) []byte {
	var namespaceTranslationDic = map[string]string{
		StrictSourceRelationship:               SourceRelationship.Value,
		StrictSourceRelationshipOfficeDocument: SourceRelationshipOfficeDocument,
		StrictSourceRelationshipChart:          SourceRelationshipChart,
		StrictSourceRelationshipComments:       SourceRelationshipComments,
		StrictSourceRelationshipImage:          SourceRelationshipImage,
		StrictNameSpaceSpreadSheet:             NameSpaceSpreadSheet.Value,
	}
	for s, n := range namespaceTranslationDic {
		content = bytesReplace(content, []byte(s), []byte(n), -1)
	}
	return content
}

// bytesReplace replace old bytes with given new.
func bytesReplace(s, old, new []byte, n int) []byte {
	if n == 0 {
		return s
	}

	if len(old) < len(new) {
		return bytes.Replace(s, old, new, n)
	}

	if n < 0 {
		n = len(s)
	}

	var wid, i, j, w int
	for i, j = 0, 0; i < len(s) && j < n; j++ {
		wid = bytes.Index(s[i:], old)
		if wid < 0 {
			break
		}

		w += copy(s[w:], s[i:i+wid])
		w += copy(s[w:], new)
		i += wid + len(old)
	}

	w += copy(s[w:], s[i:])
	return s[0:w]
}

// genSheetPasswd provides a method to generate password for worksheet
// protection by given plaintext. When an Excel sheet is being protected with
// a password, a 16-bit (two byte) long hash is generated. To verify a
// password, it is compared to the hash. Obviously, if the input data volume
// is great, numerous passwords will match the same hash. Here is the
// algorithm to create the hash value:
//
// take the ASCII values of all characters shift left the first character 1 bit,
// the second 2 bits and so on (use only the lower 15 bits and rotate all higher bits,
// the highest bit of the 16-bit value is always 0 [signed short])
// XOR all these values
// XOR the count of characters
// XOR the constant 0xCE4B
func genSheetPasswd(plaintext string) string {
	var password int64 = 0x0000
	var charPos uint = 1
	for _, v := range plaintext {
		value := int64(v) << charPos
		charPos++
		rotatedBits := value >> 15 // rotated bits beyond bit 15
		value &= 0x7fff            // first 15 bits
		password ^= (value | rotatedBits)
	}
	password ^= int64(len(plaintext))
	password ^= 0xCE4B
	return strings.ToUpper(strconv.FormatInt(password, 16))
}

// getRootElement extract root element attributes by given XML decoder.
func getRootElement(d *xml.Decoder) []xml.Attr {
	tokenIdx := 0
	for {
		token, _ := d.Token()
		if token == nil {
			break
		}
		switch startElement := token.(type) {
		case xml.StartElement:
			tokenIdx++
			if tokenIdx == 1 {
				return startElement.Attr
			}
		}
	}
	return nil
}

// genXMLNamespace generate serialized XML attributes with a multi namespace
// by given element attributes.
func genXMLNamespace(attr []xml.Attr) string {
	var rootElement string
	for _, v := range attr {
		if lastSpace := getXMLNamespace(v.Name.Space, attr); lastSpace != "" {
			if lastSpace == NameSpaceXML {
				lastSpace = "xml"
			}
			rootElement += fmt.Sprintf("%s:%s=\"%s\" ", lastSpace, v.Name.Local, v.Value)
			continue
		}
		rootElement += fmt.Sprintf("%s=\"%s\" ", v.Name.Local, v.Value)
	}
	return strings.TrimSpace(rootElement) + ">"
}

// getXMLNamespace extract XML namespace from specified element name and attributes.
func getXMLNamespace(space string, attr []xml.Attr) string {
	for _, attribute := range attr {
		if attribute.Value == space {
			return attribute.Name.Local
		}
	}
	return space
}

// replaceNameSpaceBytes provides a function to replace the XML root element
// attribute by the given component part path and XML content.
func (f *File) replaceNameSpaceBytes(path string, contentMarshal []byte) []byte {
	var oldXmlns = []byte(`xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">`)
	var newXmlns = []byte(templateNamespaceIDMap)
	if attr, ok := f.xmlAttr[path]; ok {
		newXmlns = []byte(genXMLNamespace(attr))
	}
	return bytesReplace(contentMarshal, oldXmlns, newXmlns, -1)
}

// addNameSpaces provides a function to add a XML attribute by the given
// component part path.
func (f *File) addNameSpaces(path string, ns xml.Attr) {
	exist := false
	mc := false
	ignore := -1
	if attr, ok := f.xmlAttr[path]; ok {
		for i, attribute := range attr {
			if attribute.Name.Local == ns.Name.Local && attribute.Name.Space == ns.Name.Space {
				exist = true
			}
			if attribute.Name.Local == "Ignorable" && getXMLNamespace(attribute.Name.Space, attr) == "mc" {
				ignore = i
			}
			if attribute.Name.Local == "mc" && attribute.Name.Space == "xmlns" {
				mc = true
			}
		}
	}
	if !exist {
		f.xmlAttr[path] = append(f.xmlAttr[path], ns)
		if !mc {
			f.xmlAttr[path] = append(f.xmlAttr[path], SourceRelationshipCompatibility)
		}
		if ignore == -1 {
			f.xmlAttr[path] = append(f.xmlAttr[path], xml.Attr{
				Name:  xml.Name{Local: "Ignorable", Space: "mc"},
				Value: ns.Name.Local,
			})
			return
		}
		f.setIgnorableNameSpace(path, ignore, ns)
	}
}

// setIgnorableNameSpace provides a function to set XML namespace as ignorable
// by the given attribute.
func (f *File) setIgnorableNameSpace(path string, index int, ns xml.Attr) {
	ignorableNS := []string{"c14", "cdr14", "a14", "pic14", "x14", "xdr14", "x14ac", "dsp", "mso14", "dgm14", "x15", "x12ac", "x15ac", "xr", "xr2", "xr3", "xr4", "xr5", "xr6", "xr7", "xr8", "xr9", "xr10", "xr11", "xr12", "xr13", "xr14", "xr15", "x15", "x16", "x16r2", "mo", "mx", "mv", "o", "v"}
	if inStrSlice(strings.Fields(f.xmlAttr[path][index].Value), ns.Name.Local) == -1 && inStrSlice(ignorableNS, ns.Name.Local) != -1 {
		f.xmlAttr[path][index].Value = strings.TrimSpace(fmt.Sprintf("%s %s", f.xmlAttr[path][index].Value, ns.Name.Local))
	}
}

// addSheetNameSpace add XML attribute for worksheet.
func (f *File) addSheetNameSpace(sheet string, ns xml.Attr) {
	name := f.sheetMap[trimSheetName(sheet)]
	f.addNameSpaces(name, ns)
}

// isNumeric determines whether an expression is a valid numeric type and get
// the precision for the numeric.
func isNumeric(s string) (bool, int) {
	dot := false
	p := 0
	for i, v := range s {
		if v == '.' {
			if dot {
				return false, 0
			}
			dot = true
		} else if v < '0' || v > '9' {
			if i == 0 && v == '-' {
				continue
			}
			return false, 0
		} else if dot {
			p++
		}
	}
	return true, p
}

var (
	bstrExp       = regexp.MustCompile(`_x[a-zA-Z\d]{4}_`)
	bstrEscapeExp = regexp.MustCompile(`x[a-zA-Z\d]{4}_`)
)

// bstrUnmarshal parses the binary basic string, this will trim escaped string
// literal which not permitted in an XML 1.0 document. The basic string
// variant type can store any valid Unicode character. Unicode characters
// that cannot be directly represented in XML as defined by the XML 1.0
// specification, shall be escaped using the Unicode numerical character
// representation escape character format _xHHHH_, where H represents a
// hexadecimal character in the character's value. For example: The Unicode
// character 8 is not permitted in an XML 1.0 document, so it shall be
// escaped as _x0008_. To store the literal form of an escape sequence, the
// initial underscore shall itself be escaped (i.e. stored as _x005F_). For
// example: The string literal _x0008_ would be stored as _x005F_x0008_.
func bstrUnmarshal(s string) (result string) {
	matches, l, cursor := bstrExp.FindAllStringSubmatchIndex(s, -1), len(s), 0
	for _, match := range matches {
		result += s[cursor:match[0]]
		subStr := s[match[0]:match[1]]
		if subStr == "_x005F_" {
			cursor = match[1]
			if l > match[1]+6 && !bstrEscapeExp.MatchString(s[match[1]:match[1]+6]) {
				result += subStr
				continue
			}
			result += "_"
			continue
		}
		if bstrExp.MatchString(subStr) {
			cursor = match[1]
			v, err := strconv.Unquote(`"\u` + s[match[0]+2:match[1]-1] + `"`)
			if err != nil {
				if l > match[1]+6 && bstrEscapeExp.MatchString(s[match[1]:match[1]+6]) {
					result += subStr[:6]
					cursor = match[1] + 6
					continue
				}
				result += subStr
				continue
			}
			hasRune := false
			for _, c := range v {
				if unicode.IsControl(c) {
					hasRune = true
				}
			}
			if !hasRune {
				result += v
			}
		}
	}
	if cursor < l {
		result += s[cursor:]
	}
	return result
}

// bstrMarshal encode the escaped string literal which not permitted in an XML
// 1.0 document.
func bstrMarshal(s string) (result string) {
	matches, l, cursor := bstrExp.FindAllStringSubmatchIndex(s, -1), len(s), 0
	for _, match := range matches {
		result += s[cursor:match[0]]
		subStr := s[match[0]:match[1]]
		if subStr == "_x005F_" {
			cursor = match[1]
			if match[1]+6 <= l && bstrEscapeExp.MatchString(s[match[1]:match[1]+6]) {
				_, err := strconv.Unquote(`"\u` + s[match[1]+1:match[1]+5] + `"`)
				if err == nil {
					result += subStr + "x005F" + subStr
					continue
				}
			}
			result += subStr + "x005F_"
			continue
		}
		if bstrExp.MatchString(subStr) {
			cursor = match[1]
			_, err := strconv.Unquote(`"\u` + s[match[0]+2:match[1]-1] + `"`)
			if err == nil {
				result += "_x005F" + subStr
				continue
			}
			result += subStr
		}
	}
	if cursor < l {
		result += s[cursor:]
	}
	return result
}

// Stack defined an abstract data type that serves as a collection of elements.
type Stack struct {
	list *list.List
}

// NewStack create a new stack.
func NewStack() *Stack {
	list := list.New()
	return &Stack{list}
}

// Push a value onto the top of the stack.
func (stack *Stack) Push(value interface{}) {
	stack.list.PushBack(value)
}

// Pop the top item of the stack and return it.
func (stack *Stack) Pop() interface{} {
	e := stack.list.Back()
	if e != nil {
		stack.list.Remove(e)
		return e.Value
	}
	return nil
}

// Peek view the top item on the stack.
func (stack *Stack) Peek() interface{} {
	e := stack.list.Back()
	if e != nil {
		return e.Value
	}
	return nil
}

// Len return the number of items in the stack.
func (stack *Stack) Len() int {
	return stack.list.Len()
}

// Empty the stack.
func (stack *Stack) Empty() bool {
	return stack.list.Len() == 0
}
