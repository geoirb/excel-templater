package excelize

import (
	"bytes"
	"container/list"
	"errors"
	"fmt"
	"math"
	"math/cmplx"
	"math/rand"
	"net/url"
	"reflect"
	"regexp"
	"sort"
	"strconv"
	"strings"
	"time"
	"unicode"
	"unsafe"

	"github.com/xuri/efp"
	"golang.org/x/text/language"
	"golang.org/x/text/message"
)

// Excel formula errors
const (
	formulaErrorDIV         = "#DIV/0!"
	formulaErrorNAME        = "#NAME?"
	formulaErrorNA          = "#N/A"
	formulaErrorNUM         = "#NUM!"
	formulaErrorVALUE       = "#VALUE!"
	formulaErrorREF         = "#REF!"
	formulaErrorNULL        = "#NULL"
	formulaErrorSPILL       = "#SPILL!"
	formulaErrorCALC        = "#CALC!"
	formulaErrorGETTINGDATA = "#GETTING_DATA"
)

// Numeric precision correct numeric values as legacy Excel application
// https://en.wikipedia.org/wiki/Numeric_precision_in_Microsoft_Excel In the
// top figure the fraction 1/9000 in Excel is displayed. Although this number
// has a decimal representation that is an infinite string of ones, Excel
// displays only the leading 15 figures. In the second line, the number one
// is added to the fraction, and again Excel displays only 15 figures.
const numericPrecision = 1000000000000000
const maxFinancialIterations = 128
const financialPercision = 1.0e-08

// cellRef defines the structure of a cell reference.
type cellRef struct {
	Col   int
	Row   int
	Sheet string
}

// cellRef defines the structure of a cell range.
type cellRange struct {
	From cellRef
	To   cellRef
}

// formula criteria condition enumeration.
const (
	_ byte = iota
	criteriaEq
	criteriaLe
	criteriaGe
	criteriaL
	criteriaG
	criteriaBeg
	criteriaEnd
	criteriaErr
)

// formulaCriteria defined formula criteria parser result.
type formulaCriteria struct {
	Type      byte
	Condition string
}

// ArgType is the type if formula argument type.
type ArgType byte

// Formula argument types enumeration.
const (
	ArgUnknown ArgType = iota
	ArgNumber
	ArgString
	ArgList
	ArgMatrix
	ArgError
	ArgEmpty
)

// formulaArg is the argument of a formula or function.
type formulaArg struct {
	SheetName            string
	Number               float64
	String               string
	List                 []formulaArg
	Matrix               [][]formulaArg
	Boolean              bool
	Error                string
	Type                 ArgType
	cellRefs, cellRanges *list.List
}

// Value returns a string data type of the formula argument.
func (fa formulaArg) Value() (value string) {
	switch fa.Type {
	case ArgNumber:
		if fa.Boolean {
			if fa.Number == 0 {
				return "FALSE"
			}
			return "TRUE"
		}
		return fmt.Sprintf("%g", fa.Number)
	case ArgString:
		return fa.String
	case ArgError:
		return fa.Error
	}
	return
}

// ToNumber returns a formula argument with number data type.
func (fa formulaArg) ToNumber() formulaArg {
	var n float64
	var err error
	switch fa.Type {
	case ArgString:
		n, err = strconv.ParseFloat(fa.String, 64)
		if err != nil {
			return newErrorFormulaArg(formulaErrorVALUE, err.Error())
		}
	case ArgNumber:
		n = fa.Number
	}
	return newNumberFormulaArg(n)
}

// ToBool returns a formula argument with boolean data type.
func (fa formulaArg) ToBool() formulaArg {
	var b bool
	var err error
	switch fa.Type {
	case ArgString:
		b, err = strconv.ParseBool(fa.String)
		if err != nil {
			return newErrorFormulaArg(formulaErrorVALUE, err.Error())
		}
	case ArgNumber:
		if fa.Boolean && fa.Number == 1 {
			b = true
		}
	}
	return newBoolFormulaArg(b)
}

// ToList returns a formula argument with array data type.
func (fa formulaArg) ToList() []formulaArg {
	switch fa.Type {
	case ArgMatrix:
		list := []formulaArg{}
		for _, row := range fa.Matrix {
			list = append(list, row...)
		}
		return list
	case ArgList:
		return fa.List
	case ArgNumber, ArgString, ArgError, ArgUnknown:
		return []formulaArg{fa}
	}
	return nil
}

// formulaFuncs is the type of the formula functions.
type formulaFuncs struct {
	f           *File
	sheet, cell string
}

// tokenPriority defined basic arithmetic operator priority.
var tokenPriority = map[string]int{
	"^":  5,
	"*":  4,
	"/":  4,
	"+":  3,
	"-":  3,
	"=":  2,
	"<>": 2,
	"<":  2,
	"<=": 2,
	">":  2,
	">=": 2,
	"&":  1,
}

// CalcCellValue provides a function to get calculated cell value. This
// feature is currently in working processing. Array formula, table formula
// and some other formulas are not supported currently.
//
// Supported formula functions:
//
//    ABS
//    ACOS
//    ACOSH
//    ACOT
//    ACOTH
//    AND
//    ARABIC
//    ASIN
//    ASINH
//    ATAN
//    ATAN2
//    ATANH
//    AVERAGE
//    AVERAGEA
//    BASE
//    BESSELI
//    BESSELJ
//    BESSELK
//    BESSELY
//    BIN2DEC
//    BIN2HEX
//    BIN2OCT
//    BITAND
//    BITLSHIFT
//    BITOR
//    BITRSHIFT
//    BITXOR
//    CEILING
//    CEILING.MATH
//    CEILING.PRECISE
//    CHAR
//    CHOOSE
//    CLEAN
//    CODE
//    COLUMN
//    COLUMNS
//    COMBIN
//    COMBINA
//    COMPLEX
//    CONCAT
//    CONCATENATE
//    COS
//    COSH
//    COT
//    COTH
//    COUNT
//    COUNTA
//    COUNTBLANK
//    CSC
//    CSCH
//    CUMIPMT
//    CUMPRINC
//    DATE
//    DATEDIF
//    DB
//    DDB
//    DEC2BIN
//    DEC2HEX
//    DEC2OCT
//    DECIMAL
//    DEGREES
//    DOLLARDE
//    DOLLARFR
//    EFFECT
//    ENCODEURL
//    EVEN
//    EXACT
//    EXP
//    FACT
//    FACTDOUBLE
//    FALSE
//    FIND
//    FINDB
//    FISHER
//    FISHERINV
//    FIXED
//    FLOOR
//    FLOOR.MATH
//    FLOOR.PRECISE
//    FV
//    FVSCHEDULE
//    GAMMA
//    GAMMALN
//    GCD
//    HARMEAN
//    HEX2BIN
//    HEX2DEC
//    HEX2OCT
//    HLOOKUP
//    IF
//    IFERROR
//    IMABS
//    IMAGINARY
//    IMARGUMENT
//    IMCONJUGATE
//    IMCOS
//    IMCOSH
//    IMCOT
//    IMCSC
//    IMCSCH
//    IMDIV
//    IMEXP
//    IMLN
//    IMLOG10
//    IMLOG2
//    IMPOWER
//    IMPRODUCT
//    IMREAL
//    IMSEC
//    IMSECH
//    IMSIN
//    IMSINH
//    IMSQRT
//    IMSUB
//    IMSUM
//    IMTAN
//    INT
//    IPMT
//    IRR
//    ISBLANK
//    ISERR
//    ISERROR
//    ISEVEN
//    ISNA
//    ISNONTEXT
//    ISNUMBER
//    ISODD
//    ISTEXT
//    ISO.CEILING
//    ISPMT
//    KURT
//    LARGE
//    LCM
//    LEFT
//    LEFTB
//    LEN
//    LENB
//    LN
//    LOG
//    LOG10
//    LOOKUP
//    LOWER
//    MAX
//    MDETERM
//    MEDIAN
//    MID
//    MIDB
//    MIN
//    MINA
//    MIRR
//    MOD
//    MROUND
//    MULTINOMIAL
//    MUNIT
//    N
//    NA
//    NOMINAL
//    NORM.DIST
//    NORMDIST
//    NORM.INV
//    NORMINV
//    NORM.S.DIST
//    NORMSDIST
//    NORM.S.INV
//    NORMSINV
//    NOT
//    NOW
//    NPER
//    NPV
//    OCT2BIN
//    OCT2DEC
//    OCT2HEX
//    ODD
//    OR
//    PDURATION
//    PERCENTILE.INC
//    PERCENTILE
//    PERMUT
//    PERMUTATIONA
//    PI
//    PMT
//    POISSON.DIST
//    POISSON
//    POWER
//    PPMT
//    PRODUCT
//    PROPER
//    QUARTILE
//    QUARTILE.INC
//    QUOTIENT
//    RADIANS
//    RAND
//    RANDBETWEEN
//    REPLACE
//    REPLACEB
//    REPT
//    RIGHT
//    RIGHTB
//    ROMAN
//    ROUND
//    ROUNDDOWN
//    ROUNDUP
//    ROW
//    ROWS
//    SEC
//    SECH
//    SHEET
//    SIGN
//    SIN
//    SINH
//    SKEW
//    SMALL
//    SQRT
//    SQRTPI
//    STDEV
//    STDEV.S
//    STDEVA
//    SUBSTITUTE
//    SUM
//    SUMIF
//    SUMSQ
//    T
//    TAN
//    TANH
//    TODAY
//    TRIM
//    TRUE
//    TRUNC
//    UNICHAR
//    UNICODE
//    UPPER
//    VAR.P
//    VARP
//    VLOOKUP
//
func (f *File) CalcCellValue(sheet, cell string) (result string, err error) {
	var (
		formula string
		token   efp.Token
	)
	if formula, err = f.GetCellFormula(sheet, cell); err != nil {
		return
	}
	ps := efp.ExcelParser()
	tokens := ps.Parse(formula)
	if tokens == nil {
		return
	}
	if token, err = f.evalInfixExp(sheet, cell, tokens); err != nil {
		return
	}
	result = token.TValue
	isNum, precision := isNumeric(result)
	if isNum && precision > 15 {
		num, _ := roundPrecision(result)
		result = strings.ToUpper(num)
	}
	return
}

// getPriority calculate arithmetic operator priority.
func getPriority(token efp.Token) (pri int) {
	pri = tokenPriority[token.TValue]
	if token.TValue == "-" && token.TType == efp.TokenTypeOperatorPrefix {
		pri = 6
	}
	if isBeginParenthesesToken(token) { // (
		pri = 0
	}
	return
}

// newNumberFormulaArg constructs a number formula argument.
func newNumberFormulaArg(n float64) formulaArg {
	if math.IsNaN(n) {
		return newErrorFormulaArg(formulaErrorNUM, formulaErrorNUM)
	}
	return formulaArg{Type: ArgNumber, Number: n}
}

// newStringFormulaArg constructs a string formula argument.
func newStringFormulaArg(s string) formulaArg {
	return formulaArg{Type: ArgString, String: s}
}

// newMatrixFormulaArg constructs a matrix formula argument.
func newMatrixFormulaArg(m [][]formulaArg) formulaArg {
	return formulaArg{Type: ArgMatrix, Matrix: m}
}

// newListFormulaArg create a list formula argument.
func newListFormulaArg(l []formulaArg) formulaArg {
	return formulaArg{Type: ArgList, List: l}
}

// newBoolFormulaArg constructs a boolean formula argument.
func newBoolFormulaArg(b bool) formulaArg {
	var n float64
	if b {
		n = 1
	}
	return formulaArg{Type: ArgNumber, Number: n, Boolean: true}
}

// newErrorFormulaArg create an error formula argument of a given type with a
// specified error message.
func newErrorFormulaArg(formulaError, msg string) formulaArg {
	return formulaArg{Type: ArgError, String: formulaError, Error: msg}
}

// newEmptyFormulaArg create an empty formula argument.
func newEmptyFormulaArg() formulaArg {
	return formulaArg{Type: ArgEmpty}
}

// evalInfixExp evaluate syntax analysis by given infix expression after
// lexical analysis. Evaluate an infix expression containing formulas by
// stacks:
//
//    opd  - Operand
//    opt  - Operator
//    opf  - Operation formula
//    opfd - Operand of the operation formula
//    opft - Operator of the operation formula
//    args - Arguments list of the operation formula
//
// TODO: handle subtypes: Nothing, Text, Logical, Error, Concatenation, Intersection, Union
//
func (f *File) evalInfixExp(sheet, cell string, tokens []efp.Token) (efp.Token, error) {
	var err error
	opdStack, optStack, opfStack, opfdStack, opftStack, argsStack := NewStack(), NewStack(), NewStack(), NewStack(), NewStack(), NewStack()
	for i := 0; i < len(tokens); i++ {
		token := tokens[i]

		// out of function stack
		if opfStack.Len() == 0 {
			if err = f.parseToken(sheet, token, opdStack, optStack); err != nil {
				return efp.Token{}, err
			}
		}

		// function start
		if isFunctionStartToken(token) {
			opfStack.Push(token)
			argsStack.Push(list.New().Init())
			continue
		}

		// in function stack, walk 2 token at once
		if opfStack.Len() > 0 {
			var nextToken efp.Token
			if i+1 < len(tokens) {
				nextToken = tokens[i+1]
			}

			// current token is args or range, skip next token, order required: parse reference first
			if token.TSubType == efp.TokenSubTypeRange {
				if !opftStack.Empty() {
					// parse reference: must reference at here
					result, err := f.parseReference(sheet, token.TValue)
					if err != nil {
						return efp.Token{TValue: formulaErrorNAME}, err
					}
					if result.Type != ArgString {
						return efp.Token{}, errors.New(formulaErrorVALUE)
					}
					opfdStack.Push(efp.Token{
						TType:    efp.TokenTypeOperand,
						TSubType: efp.TokenSubTypeNumber,
						TValue:   result.String,
					})
					continue
				}
				if nextToken.TType == efp.TokenTypeArgument || nextToken.TType == efp.TokenTypeFunction {
					// parse reference: reference or range at here
					refTo := f.getDefinedNameRefTo(token.TValue, sheet)
					if refTo != "" {
						token.TValue = refTo
					}
					result, err := f.parseReference(sheet, token.TValue)
					if err != nil {
						return efp.Token{TValue: formulaErrorNAME}, err
					}
					if result.Type == ArgUnknown {
						return efp.Token{}, errors.New(formulaErrorVALUE)
					}
					argsStack.Peek().(*list.List).PushBack(result)
					continue
				}
			}

			// check current token is opft
			if err = f.parseToken(sheet, token, opfdStack, opftStack); err != nil {
				return efp.Token{}, err
			}

			// current token is arg
			if token.TType == efp.TokenTypeArgument {
				for !opftStack.Empty() {
					// calculate trigger
					topOpt := opftStack.Peek().(efp.Token)
					if err := calculate(opfdStack, topOpt); err != nil {
						argsStack.Peek().(*list.List).PushFront(newErrorFormulaArg(formulaErrorVALUE, err.Error()))
					}
					opftStack.Pop()
				}
				if !opfdStack.Empty() {
					argsStack.Peek().(*list.List).PushBack(newStringFormulaArg(opfdStack.Pop().(efp.Token).TValue))
				}
				continue
			}

			// current token is logical
			if token.TType == efp.TokenTypeOperand && token.TSubType == efp.TokenSubTypeLogical {
				argsStack.Peek().(*list.List).PushBack(newStringFormulaArg(token.TValue))
			}

			if err = f.evalInfixExpFunc(sheet, cell, token, nextToken, opfStack, opdStack, opftStack, opfdStack, argsStack); err != nil {
				return efp.Token{}, err
			}
		}
	}
	for optStack.Len() != 0 {
		topOpt := optStack.Peek().(efp.Token)
		if err = calculate(opdStack, topOpt); err != nil {
			return efp.Token{}, err
		}
		optStack.Pop()
	}
	if opdStack.Len() == 0 {
		return efp.Token{}, ErrInvalidFormula
	}
	return opdStack.Peek().(efp.Token), err
}

// evalInfixExpFunc evaluate formula function in the infix expression.
func (f *File) evalInfixExpFunc(sheet, cell string, token, nextToken efp.Token, opfStack, opdStack, opftStack, opfdStack, argsStack *Stack) error {
	if !isFunctionStopToken(token) {
		return nil
	}
	// current token is function stop
	for !opftStack.Empty() {
		// calculate trigger
		topOpt := opftStack.Peek().(efp.Token)
		if err := calculate(opfdStack, topOpt); err != nil {
			return err
		}
		opftStack.Pop()
	}

	// push opfd to args
	if opfdStack.Len() > 0 {
		argsStack.Peek().(*list.List).PushBack(newStringFormulaArg(opfdStack.Pop().(efp.Token).TValue))
	}
	// call formula function to evaluate
	arg := callFuncByName(&formulaFuncs{f: f, sheet: sheet, cell: cell}, strings.NewReplacer(
		"_xlfn.", "", ".", "dot").Replace(opfStack.Peek().(efp.Token).TValue),
		[]reflect.Value{reflect.ValueOf(argsStack.Peek().(*list.List))})
	if arg.Type == ArgError && opfStack.Len() == 1 {
		return errors.New(arg.Value())
	}
	argsStack.Pop()
	opfStack.Pop()
	if opfStack.Len() > 0 { // still in function stack
		if nextToken.TType == efp.TokenTypeOperatorInfix {
			// mathematics calculate in formula function
			opfdStack.Push(efp.Token{TValue: arg.Value(), TType: efp.TokenTypeOperand, TSubType: efp.TokenSubTypeNumber})
		} else {
			argsStack.Peek().(*list.List).PushBack(arg)
		}
	} else {
		opdStack.Push(efp.Token{TValue: arg.Value(), TType: efp.TokenTypeOperand, TSubType: efp.TokenSubTypeNumber})
	}
	return nil
}

// calcPow evaluate exponentiation arithmetic operations.
func calcPow(rOpd, lOpd string, opdStack *Stack) error {
	lOpdVal, err := strconv.ParseFloat(lOpd, 64)
	if err != nil {
		return err
	}
	rOpdVal, err := strconv.ParseFloat(rOpd, 64)
	if err != nil {
		return err
	}
	result := math.Pow(lOpdVal, rOpdVal)
	opdStack.Push(efp.Token{TValue: fmt.Sprintf("%g", result), TType: efp.TokenTypeOperand, TSubType: efp.TokenSubTypeNumber})
	return nil
}

// calcEq evaluate equal arithmetic operations.
func calcEq(rOpd, lOpd string, opdStack *Stack) error {
	opdStack.Push(efp.Token{TValue: strings.ToUpper(strconv.FormatBool(rOpd == lOpd)), TType: efp.TokenTypeOperand, TSubType: efp.TokenSubTypeNumber})
	return nil
}

// calcNEq evaluate not equal arithmetic operations.
func calcNEq(rOpd, lOpd string, opdStack *Stack) error {
	opdStack.Push(efp.Token{TValue: strings.ToUpper(strconv.FormatBool(rOpd != lOpd)), TType: efp.TokenTypeOperand, TSubType: efp.TokenSubTypeNumber})
	return nil
}

// calcL evaluate less than arithmetic operations.
func calcL(rOpd, lOpd string, opdStack *Stack) error {
	lOpdVal, err := strconv.ParseFloat(lOpd, 64)
	if err != nil {
		return err
	}
	rOpdVal, err := strconv.ParseFloat(rOpd, 64)
	if err != nil {
		return err
	}
	opdStack.Push(efp.Token{TValue: strings.ToUpper(strconv.FormatBool(rOpdVal > lOpdVal)), TType: efp.TokenTypeOperand, TSubType: efp.TokenSubTypeNumber})
	return nil
}

// calcLe evaluate less than or equal arithmetic operations.
func calcLe(rOpd, lOpd string, opdStack *Stack) error {
	lOpdVal, err := strconv.ParseFloat(lOpd, 64)
	if err != nil {
		return err
	}
	rOpdVal, err := strconv.ParseFloat(rOpd, 64)
	if err != nil {
		return err
	}
	opdStack.Push(efp.Token{TValue: strings.ToUpper(strconv.FormatBool(rOpdVal >= lOpdVal)), TType: efp.TokenTypeOperand, TSubType: efp.TokenSubTypeNumber})
	return nil
}

// calcG evaluate greater than or equal arithmetic operations.
func calcG(rOpd, lOpd string, opdStack *Stack) error {
	lOpdVal, err := strconv.ParseFloat(lOpd, 64)
	if err != nil {
		return err
	}
	rOpdVal, err := strconv.ParseFloat(rOpd, 64)
	if err != nil {
		return err
	}
	opdStack.Push(efp.Token{TValue: strings.ToUpper(strconv.FormatBool(rOpdVal < lOpdVal)), TType: efp.TokenTypeOperand, TSubType: efp.TokenSubTypeNumber})
	return nil
}

// calcGe evaluate greater than or equal arithmetic operations.
func calcGe(rOpd, lOpd string, opdStack *Stack) error {
	lOpdVal, err := strconv.ParseFloat(lOpd, 64)
	if err != nil {
		return err
	}
	rOpdVal, err := strconv.ParseFloat(rOpd, 64)
	if err != nil {
		return err
	}
	opdStack.Push(efp.Token{TValue: strings.ToUpper(strconv.FormatBool(rOpdVal <= lOpdVal)), TType: efp.TokenTypeOperand, TSubType: efp.TokenSubTypeNumber})
	return nil
}

// calcSplice evaluate splice '&' operations.
func calcSplice(rOpd, lOpd string, opdStack *Stack) error {
	opdStack.Push(efp.Token{TValue: lOpd + rOpd, TType: efp.TokenTypeOperand, TSubType: efp.TokenSubTypeNumber})
	return nil
}

// calcAdd evaluate addition arithmetic operations.
func calcAdd(rOpd, lOpd string, opdStack *Stack) error {
	lOpdVal, err := strconv.ParseFloat(lOpd, 64)
	if err != nil {
		return err
	}
	rOpdVal, err := strconv.ParseFloat(rOpd, 64)
	if err != nil {
		return err
	}
	result := lOpdVal + rOpdVal
	opdStack.Push(efp.Token{TValue: fmt.Sprintf("%g", result), TType: efp.TokenTypeOperand, TSubType: efp.TokenSubTypeNumber})
	return nil
}

// calcSubtract evaluate subtraction arithmetic operations.
func calcSubtract(rOpd, lOpd string, opdStack *Stack) error {
	lOpdVal, err := strconv.ParseFloat(lOpd, 64)
	if err != nil {
		return err
	}
	rOpdVal, err := strconv.ParseFloat(rOpd, 64)
	if err != nil {
		return err
	}
	result := lOpdVal - rOpdVal
	opdStack.Push(efp.Token{TValue: fmt.Sprintf("%g", result), TType: efp.TokenTypeOperand, TSubType: efp.TokenSubTypeNumber})
	return nil
}

// calcMultiply evaluate multiplication arithmetic operations.
func calcMultiply(rOpd, lOpd string, opdStack *Stack) error {
	lOpdVal, err := strconv.ParseFloat(lOpd, 64)
	if err != nil {
		return err
	}
	rOpdVal, err := strconv.ParseFloat(rOpd, 64)
	if err != nil {
		return err
	}
	result := lOpdVal * rOpdVal
	opdStack.Push(efp.Token{TValue: fmt.Sprintf("%g", result), TType: efp.TokenTypeOperand, TSubType: efp.TokenSubTypeNumber})
	return nil
}

// calcDiv evaluate division arithmetic operations.
func calcDiv(rOpd, lOpd string, opdStack *Stack) error {
	lOpdVal, err := strconv.ParseFloat(lOpd, 64)
	if err != nil {
		return err
	}
	rOpdVal, err := strconv.ParseFloat(rOpd, 64)
	if err != nil {
		return err
	}
	result := lOpdVal / rOpdVal
	if rOpdVal == 0 {
		return errors.New(formulaErrorDIV)
	}
	opdStack.Push(efp.Token{TValue: fmt.Sprintf("%g", result), TType: efp.TokenTypeOperand, TSubType: efp.TokenSubTypeNumber})
	return nil
}

// calculate evaluate basic arithmetic operations.
func calculate(opdStack *Stack, opt efp.Token) error {
	if opt.TValue == "-" && opt.TType == efp.TokenTypeOperatorPrefix {
		if opdStack.Len() < 1 {
			return ErrInvalidFormula
		}
		opd := opdStack.Pop().(efp.Token)
		opdVal, err := strconv.ParseFloat(opd.TValue, 64)
		if err != nil {
			return err
		}
		result := 0 - opdVal
		opdStack.Push(efp.Token{TValue: fmt.Sprintf("%g", result), TType: efp.TokenTypeOperand, TSubType: efp.TokenSubTypeNumber})
	}
	tokenCalcFunc := map[string]func(rOpd, lOpd string, opdStack *Stack) error{
		"^":  calcPow,
		"*":  calcMultiply,
		"/":  calcDiv,
		"+":  calcAdd,
		"=":  calcEq,
		"<>": calcNEq,
		"<":  calcL,
		"<=": calcLe,
		">":  calcG,
		">=": calcGe,
		"&":  calcSplice,
	}
	if opt.TValue == "-" && opt.TType == efp.TokenTypeOperatorInfix {
		if opdStack.Len() < 2 {
			return ErrInvalidFormula
		}
		rOpd := opdStack.Pop().(efp.Token)
		lOpd := opdStack.Pop().(efp.Token)
		if err := calcSubtract(rOpd.TValue, lOpd.TValue, opdStack); err != nil {
			return err
		}
	}
	fn, ok := tokenCalcFunc[opt.TValue]
	if ok {
		if opdStack.Len() < 2 {
			return ErrInvalidFormula
		}
		rOpd := opdStack.Pop().(efp.Token)
		lOpd := opdStack.Pop().(efp.Token)
		if err := fn(rOpd.TValue, lOpd.TValue, opdStack); err != nil {
			return err
		}
	}
	return nil
}

// parseOperatorPrefixToken parse operator prefix token.
func (f *File) parseOperatorPrefixToken(optStack, opdStack *Stack, token efp.Token) (err error) {
	if optStack.Len() == 0 {
		optStack.Push(token)
	} else {
		tokenPriority := getPriority(token)
		topOpt := optStack.Peek().(efp.Token)
		topOptPriority := getPriority(topOpt)
		if tokenPriority > topOptPriority {
			optStack.Push(token)
		} else {
			for tokenPriority <= topOptPriority {
				optStack.Pop()
				if err = calculate(opdStack, topOpt); err != nil {
					return
				}
				if optStack.Len() > 0 {
					topOpt = optStack.Peek().(efp.Token)
					topOptPriority = getPriority(topOpt)
					continue
				}
				break
			}
			optStack.Push(token)
		}
	}
	return
}

// isFunctionStartToken determine if the token is function stop.
func isFunctionStartToken(token efp.Token) bool {
	return token.TType == efp.TokenTypeFunction && token.TSubType == efp.TokenSubTypeStart
}

// isFunctionStopToken determine if the token is function stop.
func isFunctionStopToken(token efp.Token) bool {
	return token.TType == efp.TokenTypeFunction && token.TSubType == efp.TokenSubTypeStop
}

// isBeginParenthesesToken determine if the token is begin parentheses: (.
func isBeginParenthesesToken(token efp.Token) bool {
	return token.TType == efp.TokenTypeSubexpression && token.TSubType == efp.TokenSubTypeStart
}

// isEndParenthesesToken determine if the token is end parentheses: ).
func isEndParenthesesToken(token efp.Token) bool {
	return token.TType == efp.TokenTypeSubexpression && token.TSubType == efp.TokenSubTypeStop
}

// isOperatorPrefixToken determine if the token is parse operator prefix
// token.
func isOperatorPrefixToken(token efp.Token) bool {
	_, ok := tokenPriority[token.TValue]
	return (token.TValue == "-" && token.TType == efp.TokenTypeOperatorPrefix) || (ok && token.TType == efp.TokenTypeOperatorInfix)
}

// getDefinedNameRefTo convert defined name to reference range.
func (f *File) getDefinedNameRefTo(definedNameName string, currentSheet string) (refTo string) {
	var workbookRefTo, worksheetRefTo string
	for _, definedName := range f.GetDefinedName() {
		if definedName.Name == definedNameName {
			// worksheet scope takes precedence over scope workbook when both definedNames exist
			if definedName.Scope == "Workbook" {
				workbookRefTo = definedName.RefersTo
			}
			if definedName.Scope == currentSheet {
				worksheetRefTo = definedName.RefersTo
			}
		}
	}
	refTo = workbookRefTo
	if worksheetRefTo != "" {
		refTo = worksheetRefTo
	}
	return
}

// parseToken parse basic arithmetic operator priority and evaluate based on
// operators and operands.
func (f *File) parseToken(sheet string, token efp.Token, opdStack, optStack *Stack) error {
	// parse reference: must reference at here
	if token.TSubType == efp.TokenSubTypeRange {
		refTo := f.getDefinedNameRefTo(token.TValue, sheet)
		if refTo != "" {
			token.TValue = refTo
		}
		result, err := f.parseReference(sheet, token.TValue)
		if err != nil {
			return errors.New(formulaErrorNAME)
		}
		if result.Type != ArgString {
			return errors.New(formulaErrorVALUE)
		}
		token.TValue = result.String
		token.TType = efp.TokenTypeOperand
		token.TSubType = efp.TokenSubTypeNumber
	}
	if isOperatorPrefixToken(token) {
		if err := f.parseOperatorPrefixToken(optStack, opdStack, token); err != nil {
			return err
		}
	}
	if isBeginParenthesesToken(token) { // (
		optStack.Push(token)
	}
	if isEndParenthesesToken(token) { // )
		for !isBeginParenthesesToken(optStack.Peek().(efp.Token)) { // != (
			topOpt := optStack.Peek().(efp.Token)
			if err := calculate(opdStack, topOpt); err != nil {
				return err
			}
			optStack.Pop()
		}
		optStack.Pop()
	}
	// opd
	if token.TType == efp.TokenTypeOperand && (token.TSubType == efp.TokenSubTypeNumber || token.TSubType == efp.TokenSubTypeText) {
		opdStack.Push(token)
	}
	return nil
}

// parseReference parse reference and extract values by given reference
// characters and default sheet name.
func (f *File) parseReference(sheet, reference string) (arg formulaArg, err error) {
	reference = strings.Replace(reference, "$", "", -1)
	refs, cellRanges, cellRefs := list.New(), list.New(), list.New()
	for _, ref := range strings.Split(reference, ":") {
		tokens := strings.Split(ref, "!")
		cr := cellRef{}
		if len(tokens) == 2 { // have a worksheet name
			cr.Sheet = tokens[0]
			// cast to cell coordinates
			if cr.Col, cr.Row, err = CellNameToCoordinates(tokens[1]); err != nil {
				// cast to column
				if cr.Col, err = ColumnNameToNumber(tokens[1]); err != nil {
					// cast to row
					if cr.Row, err = strconv.Atoi(tokens[1]); err != nil {
						err = newInvalidColumnNameError(tokens[1])
						return
					}
					cr.Col = TotalColumns
				}
			}
			if refs.Len() > 0 {
				e := refs.Back()
				cellRefs.PushBack(e.Value.(cellRef))
				refs.Remove(e)
			}
			refs.PushBack(cr)
			continue
		}
		// cast to cell coordinates
		if cr.Col, cr.Row, err = CellNameToCoordinates(tokens[0]); err != nil {
			// cast to column
			if cr.Col, err = ColumnNameToNumber(tokens[0]); err != nil {
				// cast to row
				if cr.Row, err = strconv.Atoi(tokens[0]); err != nil {
					err = newInvalidColumnNameError(tokens[0])
					return
				}
				cr.Col = TotalColumns
			}
			cellRanges.PushBack(cellRange{
				From: cellRef{Sheet: sheet, Col: cr.Col, Row: 1},
				To:   cellRef{Sheet: sheet, Col: cr.Col, Row: TotalRows},
			})
			cellRefs.Init()
			arg, err = f.rangeResolver(cellRefs, cellRanges)
			return
		}
		e := refs.Back()
		if e == nil {
			cr.Sheet = sheet
			refs.PushBack(cr)
			continue
		}
		cellRanges.PushBack(cellRange{
			From: e.Value.(cellRef),
			To:   cr,
		})
		refs.Remove(e)
	}
	if refs.Len() > 0 {
		e := refs.Back()
		cellRefs.PushBack(e.Value.(cellRef))
		refs.Remove(e)
	}
	arg, err = f.rangeResolver(cellRefs, cellRanges)
	return
}

// prepareValueRange prepare value range.
func prepareValueRange(cr cellRange, valueRange []int) {
	if cr.From.Row < valueRange[0] || valueRange[0] == 0 {
		valueRange[0] = cr.From.Row
	}
	if cr.From.Col < valueRange[2] || valueRange[2] == 0 {
		valueRange[2] = cr.From.Col
	}
	if cr.To.Row > valueRange[1] || valueRange[1] == 0 {
		valueRange[1] = cr.To.Row
	}
	if cr.To.Col > valueRange[3] || valueRange[3] == 0 {
		valueRange[3] = cr.To.Col
	}
}

// prepareValueRef prepare value reference.
func prepareValueRef(cr cellRef, valueRange []int) {
	if cr.Row < valueRange[0] || valueRange[0] == 0 {
		valueRange[0] = cr.Row
	}
	if cr.Col < valueRange[2] || valueRange[2] == 0 {
		valueRange[2] = cr.Col
	}
	if cr.Row > valueRange[1] || valueRange[1] == 0 {
		valueRange[1] = cr.Row
	}
	if cr.Col > valueRange[3] || valueRange[3] == 0 {
		valueRange[3] = cr.Col
	}
}

// rangeResolver extract value as string from given reference and range list.
// This function will not ignore the empty cell. For example, A1:A2:A2:B3 will
// be reference A1:B3.
func (f *File) rangeResolver(cellRefs, cellRanges *list.List) (arg formulaArg, err error) {
	arg.cellRefs, arg.cellRanges = cellRefs, cellRanges
	// value range order: from row, to row, from column, to column
	valueRange := []int{0, 0, 0, 0}
	var sheet string
	// prepare value range
	for temp := cellRanges.Front(); temp != nil; temp = temp.Next() {
		cr := temp.Value.(cellRange)
		if cr.From.Sheet != cr.To.Sheet {
			err = errors.New(formulaErrorVALUE)
		}
		rng := []int{cr.From.Col, cr.From.Row, cr.To.Col, cr.To.Row}
		_ = sortCoordinates(rng)
		cr.From.Col, cr.From.Row, cr.To.Col, cr.To.Row = rng[0], rng[1], rng[2], rng[3]
		prepareValueRange(cr, valueRange)
		if cr.From.Sheet != "" {
			sheet = cr.From.Sheet
		}
	}
	for temp := cellRefs.Front(); temp != nil; temp = temp.Next() {
		cr := temp.Value.(cellRef)
		if cr.Sheet != "" {
			sheet = cr.Sheet
		}
		prepareValueRef(cr, valueRange)
	}
	// extract value from ranges
	if cellRanges.Len() > 0 {
		arg.Type = ArgMatrix
		for row := valueRange[0]; row <= valueRange[1]; row++ {
			var matrixRow = []formulaArg{}
			for col := valueRange[2]; col <= valueRange[3]; col++ {
				var cell, value string
				if cell, err = CoordinatesToCellName(col, row); err != nil {
					return
				}
				if value, err = f.GetCellValue(sheet, cell); err != nil {
					return
				}
				matrixRow = append(matrixRow, formulaArg{
					String: value,
					Type:   ArgString,
				})
			}
			arg.Matrix = append(arg.Matrix, matrixRow)
		}
		return
	}
	// extract value from references
	for temp := cellRefs.Front(); temp != nil; temp = temp.Next() {
		cr := temp.Value.(cellRef)
		var cell string
		if cell, err = CoordinatesToCellName(cr.Col, cr.Row); err != nil {
			return
		}
		if arg.String, err = f.GetCellValue(cr.Sheet, cell); err != nil {
			return
		}
		arg.Type = ArgString
	}
	return
}

// callFuncByName calls the no error or only error return function with
// reflect by given receiver, name and parameters.
func callFuncByName(receiver interface{}, name string, params []reflect.Value) (arg formulaArg) {
	function := reflect.ValueOf(receiver).MethodByName(name)
	if function.IsValid() {
		rt := function.Call(params)
		if len(rt) == 0 {
			return
		}
		arg = rt[0].Interface().(formulaArg)
		return
	}
	return newErrorFormulaArg(formulaErrorVALUE, fmt.Sprintf("not support %s function", name))
}

// formulaCriteriaParser parse formula criteria.
func formulaCriteriaParser(exp string) (fc *formulaCriteria) {
	fc = &formulaCriteria{}
	if exp == "" {
		return
	}
	if match := regexp.MustCompile(`^([0-9]+)$`).FindStringSubmatch(exp); len(match) > 1 {
		fc.Type, fc.Condition = criteriaEq, match[1]
		return
	}
	if match := regexp.MustCompile(`^=(.*)$`).FindStringSubmatch(exp); len(match) > 1 {
		fc.Type, fc.Condition = criteriaEq, match[1]
		return
	}
	if match := regexp.MustCompile(`^<=(.*)$`).FindStringSubmatch(exp); len(match) > 1 {
		fc.Type, fc.Condition = criteriaLe, match[1]
		return
	}
	if match := regexp.MustCompile(`^>=(.*)$`).FindStringSubmatch(exp); len(match) > 1 {
		fc.Type, fc.Condition = criteriaGe, match[1]
		return
	}
	if match := regexp.MustCompile(`^<(.*)$`).FindStringSubmatch(exp); len(match) > 1 {
		fc.Type, fc.Condition = criteriaL, match[1]
		return
	}
	if match := regexp.MustCompile(`^>(.*)$`).FindStringSubmatch(exp); len(match) > 1 {
		fc.Type, fc.Condition = criteriaG, match[1]
		return
	}
	if strings.Contains(exp, "*") {
		if strings.HasPrefix(exp, "*") {
			fc.Type, fc.Condition = criteriaEnd, strings.TrimPrefix(exp, "*")
		}
		if strings.HasSuffix(exp, "*") {
			fc.Type, fc.Condition = criteriaBeg, strings.TrimSuffix(exp, "*")
		}
		return
	}
	fc.Type, fc.Condition = criteriaEq, exp
	return
}

// formulaCriteriaEval evaluate formula criteria expression.
func formulaCriteriaEval(val string, criteria *formulaCriteria) (result bool, err error) {
	var value, expected float64
	var e error
	var prepareValue = func(val, cond string) (value float64, expected float64, err error) {
		if value, err = strconv.ParseFloat(val, 64); err != nil {
			return
		}
		if expected, err = strconv.ParseFloat(criteria.Condition, 64); err != nil {
			return
		}
		return
	}
	switch criteria.Type {
	case criteriaEq:
		return val == criteria.Condition, err
	case criteriaLe:
		value, expected, e = prepareValue(val, criteria.Condition)
		return value <= expected && e == nil, err
	case criteriaGe:
		value, expected, e = prepareValue(val, criteria.Condition)
		return value >= expected && e == nil, err
	case criteriaL:
		value, expected, e = prepareValue(val, criteria.Condition)
		return value < expected && e == nil, err
	case criteriaG:
		value, expected, e = prepareValue(val, criteria.Condition)
		return value > expected && e == nil, err
	case criteriaBeg:
		return strings.HasPrefix(val, criteria.Condition), err
	case criteriaEnd:
		return strings.HasSuffix(val, criteria.Condition), err
	}
	return
}

// Engineering Functions

// BESSELI function the modified Bessel function, which is equivalent to the
// Bessel function evaluated for purely imaginary arguments. The syntax of
// the Besseli function is:
//
//    BESSELI(x,n)
//
func (fn *formulaFuncs) BESSELI(argsList *list.List) formulaArg {
	if argsList.Len() != 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "BESSELI requires 2 numeric arguments")
	}
	return fn.bassel(argsList, true)
}

// BESSELJ function returns the Bessel function, Jn(x), for a specified order
// and value of x. The syntax of the function is:
//
//    BESSELJ(x,n)
//
func (fn *formulaFuncs) BESSELJ(argsList *list.List) formulaArg {
	if argsList.Len() != 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "BESSELJ requires 2 numeric arguments")
	}
	return fn.bassel(argsList, false)
}

// bassel is an implementation of the formula function BESSELI and BESSELJ.
func (fn *formulaFuncs) bassel(argsList *list.List, modfied bool) formulaArg {
	x, n := argsList.Front().Value.(formulaArg).ToNumber(), argsList.Back().Value.(formulaArg).ToNumber()
	if x.Type != ArgNumber {
		return x
	}
	if n.Type != ArgNumber {
		return n
	}
	max, x1 := 100, x.Number*0.5
	x2 := x1 * x1
	x1 = math.Pow(x1, n.Number)
	n1, n2, n3, n4, add := fact(n.Number), 1.0, 0.0, n.Number, false
	result := x1 / n1
	t := result * 0.9
	for result != t && max != 0 {
		x1 *= x2
		n3++
		n1 *= n3
		n4++
		n2 *= n4
		t = result
		if modfied || add {
			result += (x1 / n1 / n2)
		} else {
			result -= (x1 / n1 / n2)
		}
		max--
		add = !add
	}
	return newNumberFormulaArg(result)
}

// BESSELK function calculates the modified Bessel functions, Kn(x), which are
// also known as the hyperbolic Bessel Functions. These are the equivalent of
// the Bessel functions, evaluated for purely imaginary arguments. The syntax
// of the function is:
//
//    BESSELK(x,n)
//
func (fn *formulaFuncs) BESSELK(argsList *list.List) formulaArg {
	if argsList.Len() != 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "BESSELK requires 2 numeric arguments")
	}
	x, n := argsList.Front().Value.(formulaArg).ToNumber(), argsList.Back().Value.(formulaArg).ToNumber()
	if x.Type != ArgNumber {
		return x
	}
	if n.Type != ArgNumber {
		return n
	}
	if x.Number <= 0 || n.Number < 0 {
		return newErrorFormulaArg(formulaErrorNUM, formulaErrorNUM)
	}
	var result float64
	switch math.Floor(n.Number) {
	case 0:
		result = fn.besselK0(x)
	case 1:
		result = fn.besselK1(x)
	default:
		result = fn.besselK2(x, n)
	}
	return newNumberFormulaArg(result)
}

// besselK0 is an implementation of the formula function BESSELK.
func (fn *formulaFuncs) besselK0(x formulaArg) float64 {
	var y float64
	if x.Number <= 2 {
		n2 := x.Number * 0.5
		y = n2 * n2
		args := list.New()
		args.PushBack(x)
		args.PushBack(newNumberFormulaArg(0))
		return -math.Log(n2)*fn.BESSELI(args).Number +
			(-0.57721566 + y*(0.42278420+y*(0.23069756+y*(0.3488590e-1+y*(0.262698e-2+y*
				(0.10750e-3+y*0.74e-5))))))
	}
	y = 2 / x.Number
	return math.Exp(-x.Number) / math.Sqrt(x.Number) *
		(1.25331414 + y*(-0.7832358e-1+y*(0.2189568e-1+y*(-0.1062446e-1+y*
			(0.587872e-2+y*(-0.251540e-2+y*0.53208e-3))))))
}

// besselK1 is an implementation of the formula function BESSELK.
func (fn *formulaFuncs) besselK1(x formulaArg) float64 {
	var n2, y float64
	if x.Number <= 2 {
		n2 = x.Number * 0.5
		y = n2 * n2
		args := list.New()
		args.PushBack(x)
		args.PushBack(newNumberFormulaArg(1))
		return math.Log(n2)*fn.BESSELI(args).Number +
			(1+y*(0.15443144+y*(-0.67278579+y*(-0.18156897+y*(-0.1919402e-1+y*(-0.110404e-2+y*(-0.4686e-4)))))))/x.Number
	}
	y = 2 / x.Number
	return math.Exp(-x.Number) / math.Sqrt(x.Number) *
		(1.25331414 + y*(0.23498619+y*(-0.3655620e-1+y*(0.1504268e-1+y*(-0.780353e-2+y*
			(0.325614e-2+y*(-0.68245e-3)))))))
}

// besselK2 is an implementation of the formula function BESSELK.
func (fn *formulaFuncs) besselK2(x, n formulaArg) float64 {
	tox, bkm, bk, bkp := 2/x.Number, fn.besselK0(x), fn.besselK1(x), 0.0
	for i := 1.0; i < n.Number; i++ {
		bkp = bkm + i*tox*bk
		bkm = bk
		bk = bkp
	}
	return bk
}

// BESSELY function returns the Bessel function, Yn(x), (also known as the
// Weber function or the Neumann function), for a specified order and value
// of x. The syntax of the function is:
//
//    BESSELY(x,n)
//
func (fn *formulaFuncs) BESSELY(argsList *list.List) formulaArg {
	if argsList.Len() != 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "BESSELY requires 2 numeric arguments")
	}
	x, n := argsList.Front().Value.(formulaArg).ToNumber(), argsList.Back().Value.(formulaArg).ToNumber()
	if x.Type != ArgNumber {
		return x
	}
	if n.Type != ArgNumber {
		return n
	}
	if x.Number <= 0 || n.Number < 0 {
		return newErrorFormulaArg(formulaErrorNUM, formulaErrorNUM)
	}
	var result float64
	switch math.Floor(n.Number) {
	case 0:
		result = fn.besselY0(x)
	case 1:
		result = fn.besselY1(x)
	default:
		result = fn.besselY2(x, n)
	}
	return newNumberFormulaArg(result)
}

// besselY0 is an implementation of the formula function BESSELY.
func (fn *formulaFuncs) besselY0(x formulaArg) float64 {
	var y float64
	if x.Number < 8 {
		y = x.Number * x.Number
		f1 := -2957821389.0 + y*(7062834065.0+y*(-512359803.6+y*(10879881.29+y*
			(-86327.92757+y*228.4622733))))
		f2 := 40076544269.0 + y*(745249964.8+y*(7189466.438+y*
			(47447.26470+y*(226.1030244+y))))
		args := list.New()
		args.PushBack(x)
		args.PushBack(newNumberFormulaArg(0))
		return f1/f2 + 0.636619772*fn.BESSELJ(args).Number*math.Log(x.Number)
	}
	z := 8.0 / x.Number
	y = z * z
	xx := x.Number - 0.785398164
	f1 := 1 + y*(-0.1098628627e-2+y*(0.2734510407e-4+y*(-0.2073370639e-5+y*0.2093887211e-6)))
	f2 := -0.1562499995e-1 + y*(0.1430488765e-3+y*(-0.6911147651e-5+y*(0.7621095161e-6+y*
		(-0.934945152e-7))))
	return math.Sqrt(0.636619772/x.Number) * (math.Sin(xx)*f1 + z*math.Cos(xx)*f2)
}

// besselY1 is an implementation of the formula function BESSELY.
func (fn *formulaFuncs) besselY1(x formulaArg) float64 {
	if x.Number < 8 {
		y := x.Number * x.Number
		f1 := x.Number * (-0.4900604943e13 + y*(0.1275274390e13+y*(-0.5153438139e11+y*
			(0.7349264551e9+y*(-0.4237922726e7+y*0.8511937935e4)))))
		f2 := 0.2499580570e14 + y*(0.4244419664e12+y*(0.3733650367e10+y*(0.2245904002e8+y*
			(0.1020426050e6+y*(0.3549632885e3+y)))))
		args := list.New()
		args.PushBack(x)
		args.PushBack(newNumberFormulaArg(1))
		return f1/f2 + 0.636619772*(fn.BESSELJ(args).Number*math.Log(x.Number)-1/x.Number)
	}
	return math.Sqrt(0.636619772/x.Number) * math.Sin(x.Number-2.356194491)
}

// besselY2 is an implementation of the formula function BESSELY.
func (fn *formulaFuncs) besselY2(x, n formulaArg) float64 {
	tox, bym, by, byp := 2/x.Number, fn.besselY0(x), fn.besselY1(x), 0.0
	for i := 1.0; i < n.Number; i++ {
		byp = i*tox*by - bym
		bym = by
		by = byp
	}
	return by
}

// BIN2DEC function converts a Binary (a base-2 number) into a decimal number.
// The syntax of the function is:
//
//    BIN2DEC(number)
//
func (fn *formulaFuncs) BIN2DEC(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "BIN2DEC requires 1 numeric argument")
	}
	token := argsList.Front().Value.(formulaArg)
	number := token.ToNumber()
	if number.Type != ArgNumber {
		return newErrorFormulaArg(formulaErrorVALUE, number.Error)
	}
	return fn.bin2dec(token.Value())
}

// BIN2HEX function converts a Binary (Base 2) number into a Hexadecimal
// (Base 16) number. The syntax of the function is:
//
//    BIN2HEX(number,[places])
//
func (fn *formulaFuncs) BIN2HEX(argsList *list.List) formulaArg {
	if argsList.Len() < 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "BIN2HEX requires at least 1 argument")
	}
	if argsList.Len() > 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "BIN2HEX allows at most 2 arguments")
	}
	token := argsList.Front().Value.(formulaArg)
	number := token.ToNumber()
	if number.Type != ArgNumber {
		return newErrorFormulaArg(formulaErrorVALUE, number.Error)
	}
	decimal, newList := fn.bin2dec(token.Value()), list.New()
	if decimal.Type != ArgNumber {
		return decimal
	}
	newList.PushBack(decimal)
	if argsList.Len() == 2 {
		newList.PushBack(argsList.Back().Value.(formulaArg))
	}
	return fn.dec2x("BIN2HEX", newList)
}

// BIN2OCT function converts a Binary (Base 2) number into an Octal (Base 8)
// number. The syntax of the function is:
//
//    BIN2OCT(number,[places])
//
func (fn *formulaFuncs) BIN2OCT(argsList *list.List) formulaArg {
	if argsList.Len() < 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "BIN2OCT requires at least 1 argument")
	}
	if argsList.Len() > 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "BIN2OCT allows at most 2 arguments")
	}
	token := argsList.Front().Value.(formulaArg)
	number := token.ToNumber()
	if number.Type != ArgNumber {
		return newErrorFormulaArg(formulaErrorVALUE, number.Error)
	}
	decimal, newList := fn.bin2dec(token.Value()), list.New()
	if decimal.Type != ArgNumber {
		return decimal
	}
	newList.PushBack(decimal)
	if argsList.Len() == 2 {
		newList.PushBack(argsList.Back().Value.(formulaArg))
	}
	return fn.dec2x("BIN2OCT", newList)
}

// bin2dec is an implementation of the formula function BIN2DEC.
func (fn *formulaFuncs) bin2dec(number string) formulaArg {
	decimal, length := 0.0, len(number)
	for i := length; i > 0; i-- {
		s := string(number[length-i])
		if i == 10 && s == "1" {
			decimal += math.Pow(-2.0, float64(i-1))
			continue
		}
		if s == "1" {
			decimal += math.Pow(2.0, float64(i-1))
			continue
		}
		if s != "0" {
			return newErrorFormulaArg(formulaErrorNUM, formulaErrorNUM)
		}
	}
	return newNumberFormulaArg(decimal)
}

// BITAND function returns the bitwise 'AND' for two supplied integers. The
// syntax of the function is:
//
//    BITAND(number1,number2)
//
func (fn *formulaFuncs) BITAND(argsList *list.List) formulaArg {
	return fn.bitwise("BITAND", argsList)
}

// BITLSHIFT function returns a supplied integer, shifted left by a specified
// number of bits. The syntax of the function is:
//
//    BITLSHIFT(number1,shift_amount)
//
func (fn *formulaFuncs) BITLSHIFT(argsList *list.List) formulaArg {
	return fn.bitwise("BITLSHIFT", argsList)
}

// BITOR function returns the bitwise 'OR' for two supplied integers. The
// syntax of the function is:
//
//    BITOR(number1,number2)
//
func (fn *formulaFuncs) BITOR(argsList *list.List) formulaArg {
	return fn.bitwise("BITOR", argsList)
}

// BITRSHIFT function returns a supplied integer, shifted right by a specified
// number of bits. The syntax of the function is:
//
//    BITRSHIFT(number1,shift_amount)
//
func (fn *formulaFuncs) BITRSHIFT(argsList *list.List) formulaArg {
	return fn.bitwise("BITRSHIFT", argsList)
}

// BITXOR function returns the bitwise 'XOR' (exclusive 'OR') for two supplied
// integers. The syntax of the function is:
//
//    BITXOR(number1,number2)
//
func (fn *formulaFuncs) BITXOR(argsList *list.List) formulaArg {
	return fn.bitwise("BITXOR", argsList)
}

// bitwise is an implementation of the formula function BITAND, BITLSHIFT,
// BITOR, BITRSHIFT and BITXOR.
func (fn *formulaFuncs) bitwise(name string, argsList *list.List) formulaArg {
	if argsList.Len() != 2 {
		return newErrorFormulaArg(formulaErrorVALUE, fmt.Sprintf("%s requires 2 numeric arguments", name))
	}
	num1, num2 := argsList.Front().Value.(formulaArg).ToNumber(), argsList.Back().Value.(formulaArg).ToNumber()
	if num1.Type != ArgNumber || num2.Type != ArgNumber {
		return newErrorFormulaArg(formulaErrorNUM, formulaErrorNUM)
	}
	max := math.Pow(2, 48) - 1
	if num1.Number < 0 || num1.Number > max || num2.Number < 0 || num2.Number > max {
		return newErrorFormulaArg(formulaErrorNUM, formulaErrorNUM)
	}
	bitwiseFuncMap := map[string]func(a, b int) int{
		"BITAND":    func(a, b int) int { return a & b },
		"BITLSHIFT": func(a, b int) int { return a << uint(b) },
		"BITOR":     func(a, b int) int { return a | b },
		"BITRSHIFT": func(a, b int) int { return a >> uint(b) },
		"BITXOR":    func(a, b int) int { return a ^ b },
	}
	bitwiseFunc := bitwiseFuncMap[name]
	return newNumberFormulaArg(float64(bitwiseFunc(int(num1.Number), int(num2.Number))))
}

// COMPLEX function takes two arguments, representing the real and the
// imaginary coefficients of a complex number, and from these, creates a
// complex number. The syntax of the function is:
//
//    COMPLEX(real_num,i_num,[suffix])
//
func (fn *formulaFuncs) COMPLEX(argsList *list.List) formulaArg {
	if argsList.Len() < 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "COMPLEX requires at least 2 arguments")
	}
	if argsList.Len() > 3 {
		return newErrorFormulaArg(formulaErrorVALUE, "COMPLEX allows at most 3 arguments")
	}
	real, i, suffix := argsList.Front().Value.(formulaArg).ToNumber(), argsList.Front().Next().Value.(formulaArg).ToNumber(), "i"
	if real.Type != ArgNumber {
		return real
	}
	if i.Type != ArgNumber {
		return i
	}
	if argsList.Len() == 3 {
		if suffix = strings.ToLower(argsList.Back().Value.(formulaArg).Value()); suffix != "i" && suffix != "j" {
			return newErrorFormulaArg(formulaErrorVALUE, formulaErrorVALUE)
		}
	}
	return newStringFormulaArg(cmplx2str(fmt.Sprint(complex(real.Number, i.Number)), suffix))
}

// cmplx2str replace complex number string characters.
func cmplx2str(c, suffix string) string {
	if c == "(0+0i)" || c == "(-0+0i)" || c == "(0-0i)" || c == "(-0-0i)" {
		return "0"
	}
	c = strings.TrimPrefix(c, "(")
	c = strings.TrimPrefix(c, "+0+")
	c = strings.TrimPrefix(c, "-0+")
	c = strings.TrimSuffix(c, ")")
	c = strings.TrimPrefix(c, "0+")
	if strings.HasPrefix(c, "0-") {
		c = "-" + strings.TrimPrefix(c, "0-")
	}
	c = strings.TrimPrefix(c, "0+")
	c = strings.TrimSuffix(c, "+0i")
	c = strings.TrimSuffix(c, "-0i")
	c = strings.NewReplacer("+1i", "+i", "-1i", "-i").Replace(c)
	c = strings.Replace(c, "i", suffix, -1)
	return c
}

// str2cmplx convert complex number string characters.
func str2cmplx(c string) string {
	c = strings.Replace(c, "j", "i", -1)
	if c == "i" {
		c = "1i"
	}
	c = strings.NewReplacer("+i", "+1i", "-i", "-1i").Replace(c)
	return c
}

// DEC2BIN function converts a decimal number into a Binary (Base 2) number.
// The syntax of the function is:
//
//    DEC2BIN(number,[places])
//
func (fn *formulaFuncs) DEC2BIN(argsList *list.List) formulaArg {
	return fn.dec2x("DEC2BIN", argsList)
}

// DEC2HEX function converts a decimal number into a Hexadecimal (Base 16)
// number. The syntax of the function is:
//
//    DEC2HEX(number,[places])
//
func (fn *formulaFuncs) DEC2HEX(argsList *list.List) formulaArg {
	return fn.dec2x("DEC2HEX", argsList)
}

// DEC2OCT function converts a decimal number into an Octal (Base 8) number.
// The syntax of the function is:
//
//    DEC2OCT(number,[places])
//
func (fn *formulaFuncs) DEC2OCT(argsList *list.List) formulaArg {
	return fn.dec2x("DEC2OCT", argsList)
}

// dec2x is an implementation of the formula function DEC2BIN, DEC2HEX and
// DEC2OCT.
func (fn *formulaFuncs) dec2x(name string, argsList *list.List) formulaArg {
	if argsList.Len() < 1 {
		return newErrorFormulaArg(formulaErrorVALUE, fmt.Sprintf("%s requires at least 1 argument", name))
	}
	if argsList.Len() > 2 {
		return newErrorFormulaArg(formulaErrorVALUE, fmt.Sprintf("%s allows at most 2 arguments", name))
	}
	decimal := argsList.Front().Value.(formulaArg).ToNumber()
	if decimal.Type != ArgNumber {
		return newErrorFormulaArg(formulaErrorVALUE, decimal.Error)
	}
	maxLimitMap := map[string]float64{
		"DEC2BIN": 511,
		"HEX2BIN": 511,
		"OCT2BIN": 511,
		"BIN2HEX": 549755813887,
		"DEC2HEX": 549755813887,
		"OCT2HEX": 549755813887,
		"BIN2OCT": 536870911,
		"DEC2OCT": 536870911,
		"HEX2OCT": 536870911,
	}
	minLimitMap := map[string]float64{
		"DEC2BIN": -512,
		"HEX2BIN": -512,
		"OCT2BIN": -512,
		"BIN2HEX": -549755813888,
		"DEC2HEX": -549755813888,
		"OCT2HEX": -549755813888,
		"BIN2OCT": -536870912,
		"DEC2OCT": -536870912,
		"HEX2OCT": -536870912,
	}
	baseMap := map[string]int{
		"DEC2BIN": 2,
		"HEX2BIN": 2,
		"OCT2BIN": 2,
		"BIN2HEX": 16,
		"DEC2HEX": 16,
		"OCT2HEX": 16,
		"BIN2OCT": 8,
		"DEC2OCT": 8,
		"HEX2OCT": 8,
	}
	maxLimit, minLimit := maxLimitMap[name], minLimitMap[name]
	base := baseMap[name]
	if decimal.Number < minLimit || decimal.Number > maxLimit {
		return newErrorFormulaArg(formulaErrorNUM, formulaErrorNUM)
	}
	n := int64(decimal.Number)
	binary := strconv.FormatUint(*(*uint64)(unsafe.Pointer(&n)), base)
	if argsList.Len() == 2 {
		places := argsList.Back().Value.(formulaArg).ToNumber()
		if places.Type != ArgNumber {
			return newErrorFormulaArg(formulaErrorVALUE, places.Error)
		}
		binaryPlaces := len(binary)
		if places.Number < 0 || places.Number > 10 || binaryPlaces > int(places.Number) {
			return newErrorFormulaArg(formulaErrorNUM, formulaErrorNUM)
		}
		return newStringFormulaArg(strings.ToUpper(fmt.Sprintf("%s%s", strings.Repeat("0", int(places.Number)-binaryPlaces), binary)))
	}
	if decimal.Number < 0 && len(binary) > 10 {
		return newStringFormulaArg(strings.ToUpper(binary[len(binary)-10:]))
	}
	return newStringFormulaArg(strings.ToUpper(binary))
}

// HEX2BIN function converts a Hexadecimal (Base 16) number into a Binary
// (Base 2) number. The syntax of the function is:
//
//    HEX2BIN(number,[places])
//
func (fn *formulaFuncs) HEX2BIN(argsList *list.List) formulaArg {
	if argsList.Len() < 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "HEX2BIN requires at least 1 argument")
	}
	if argsList.Len() > 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "HEX2BIN allows at most 2 arguments")
	}
	decimal, newList := fn.hex2dec(argsList.Front().Value.(formulaArg).Value()), list.New()
	if decimal.Type != ArgNumber {
		return decimal
	}
	newList.PushBack(decimal)
	if argsList.Len() == 2 {
		newList.PushBack(argsList.Back().Value.(formulaArg))
	}
	return fn.dec2x("HEX2BIN", newList)
}

// HEX2DEC function converts a hexadecimal (a base-16 number) into a decimal
// number. The syntax of the function is:
//
//    HEX2DEC(number)
//
func (fn *formulaFuncs) HEX2DEC(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "HEX2DEC requires 1 numeric argument")
	}
	return fn.hex2dec(argsList.Front().Value.(formulaArg).Value())
}

// HEX2OCT function converts a Hexadecimal (Base 16) number into an Octal
// (Base 8) number. The syntax of the function is:
//
//    HEX2OCT(number,[places])
//
func (fn *formulaFuncs) HEX2OCT(argsList *list.List) formulaArg {
	if argsList.Len() < 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "HEX2OCT requires at least 1 argument")
	}
	if argsList.Len() > 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "HEX2OCT allows at most 2 arguments")
	}
	decimal, newList := fn.hex2dec(argsList.Front().Value.(formulaArg).Value()), list.New()
	if decimal.Type != ArgNumber {
		return decimal
	}
	newList.PushBack(decimal)
	if argsList.Len() == 2 {
		newList.PushBack(argsList.Back().Value.(formulaArg))
	}
	return fn.dec2x("HEX2OCT", newList)
}

// hex2dec is an implementation of the formula function HEX2DEC.
func (fn *formulaFuncs) hex2dec(number string) formulaArg {
	decimal, length := 0.0, len(number)
	for i := length; i > 0; i-- {
		num, err := strconv.ParseInt(string(number[length-i]), 16, 64)
		if err != nil {
			return newErrorFormulaArg(formulaErrorNUM, err.Error())
		}
		if i == 10 && string(number[length-i]) == "F" {
			decimal += math.Pow(-16.0, float64(i-1))
			continue
		}
		decimal += float64(num) * math.Pow(16.0, float64(i-1))
	}
	return newNumberFormulaArg(decimal)
}

// IMABS function returns the absolute value (the modulus) of a complex
// number. The syntax of the function is:
//
//    IMABS(inumber)
//
func (fn *formulaFuncs) IMABS(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "IMABS requires 1 argument")
	}
	inumber, err := strconv.ParseComplex(str2cmplx(argsList.Front().Value.(formulaArg).Value()), 128)
	if err != nil {
		return newErrorFormulaArg(formulaErrorNUM, err.Error())
	}
	return newNumberFormulaArg(cmplx.Abs(inumber))
}

// IMAGINARY function returns the imaginary coefficient of a supplied complex
// number. The syntax of the function is:
//
//    IMAGINARY(inumber)
//
func (fn *formulaFuncs) IMAGINARY(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "IMAGINARY requires 1 argument")
	}
	inumber, err := strconv.ParseComplex(str2cmplx(argsList.Front().Value.(formulaArg).Value()), 128)
	if err != nil {
		return newErrorFormulaArg(formulaErrorNUM, err.Error())
	}
	return newNumberFormulaArg(imag(inumber))
}

// IMARGUMENT function returns the phase (also called the argument) of a
// supplied complex number. The syntax of the function is:
//
//    IMARGUMENT(inumber)
//
func (fn *formulaFuncs) IMARGUMENT(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "IMARGUMENT requires 1 argument")
	}
	inumber, err := strconv.ParseComplex(str2cmplx(argsList.Front().Value.(formulaArg).Value()), 128)
	if err != nil {
		return newErrorFormulaArg(formulaErrorNUM, err.Error())
	}
	return newNumberFormulaArg(cmplx.Phase(inumber))
}

// IMCONJUGATE function returns the complex conjugate of a supplied complex
// number. The syntax of the function is:
//
//    IMCONJUGATE(inumber)
//
func (fn *formulaFuncs) IMCONJUGATE(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "IMCONJUGATE requires 1 argument")
	}
	inumber, err := strconv.ParseComplex(str2cmplx(argsList.Front().Value.(formulaArg).Value()), 128)
	if err != nil {
		return newErrorFormulaArg(formulaErrorNUM, err.Error())
	}
	return newStringFormulaArg(cmplx2str(fmt.Sprint(cmplx.Conj(inumber)), "i"))
}

// IMCOS function returns the cosine of a supplied complex number. The syntax
// of the function is:
//
//    IMCOS(inumber)
//
func (fn *formulaFuncs) IMCOS(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "IMCOS requires 1 argument")
	}
	inumber, err := strconv.ParseComplex(str2cmplx(argsList.Front().Value.(formulaArg).Value()), 128)
	if err != nil {
		return newErrorFormulaArg(formulaErrorNUM, err.Error())
	}
	return newStringFormulaArg(cmplx2str(fmt.Sprint(cmplx.Cos(inumber)), "i"))
}

// IMCOSH function returns the hyperbolic cosine of a supplied complex number. The syntax
// of the function is:
//
//    IMCOSH(inumber)
//
func (fn *formulaFuncs) IMCOSH(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "IMCOSH requires 1 argument")
	}
	inumber, err := strconv.ParseComplex(str2cmplx(argsList.Front().Value.(formulaArg).Value()), 128)
	if err != nil {
		return newErrorFormulaArg(formulaErrorNUM, err.Error())
	}
	return newStringFormulaArg(cmplx2str(fmt.Sprint(cmplx.Cosh(inumber)), "i"))
}

// IMCOT function returns the cotangent of a supplied complex number. The syntax
// of the function is:
//
//    IMCOT(inumber)
//
func (fn *formulaFuncs) IMCOT(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "IMCOT requires 1 argument")
	}
	inumber, err := strconv.ParseComplex(str2cmplx(argsList.Front().Value.(formulaArg).Value()), 128)
	if err != nil {
		return newErrorFormulaArg(formulaErrorNUM, err.Error())
	}
	return newStringFormulaArg(cmplx2str(fmt.Sprint(cmplx.Cot(inumber)), "i"))
}

// IMCSC function returns the cosecant of a supplied complex number. The syntax
// of the function is:
//
//    IMCSC(inumber)
//
func (fn *formulaFuncs) IMCSC(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "IMCSC requires 1 argument")
	}
	inumber, err := strconv.ParseComplex(str2cmplx(argsList.Front().Value.(formulaArg).Value()), 128)
	if err != nil {
		return newErrorFormulaArg(formulaErrorNUM, err.Error())
	}
	num := 1 / cmplx.Sin(inumber)
	if cmplx.IsInf(num) {
		return newErrorFormulaArg(formulaErrorNUM, formulaErrorNUM)
	}
	return newStringFormulaArg(cmplx2str(fmt.Sprint(num), "i"))
}

// IMCSCH function returns the hyperbolic cosecant of a supplied complex
// number. The syntax of the function is:
//
//    IMCSCH(inumber)
//
func (fn *formulaFuncs) IMCSCH(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "IMCSCH requires 1 argument")
	}
	inumber, err := strconv.ParseComplex(str2cmplx(argsList.Front().Value.(formulaArg).Value()), 128)
	if err != nil {
		return newErrorFormulaArg(formulaErrorNUM, err.Error())
	}
	num := 1 / cmplx.Sinh(inumber)
	if cmplx.IsInf(num) {
		return newErrorFormulaArg(formulaErrorNUM, formulaErrorNUM)
	}
	return newStringFormulaArg(cmplx2str(fmt.Sprint(num), "i"))
}

// IMDIV function calculates the quotient of two complex numbers (i.e. divides
// one complex number by another). The syntax of the function is:
//
//    IMDIV(inumber1,inumber2)
//
func (fn *formulaFuncs) IMDIV(argsList *list.List) formulaArg {
	if argsList.Len() != 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "IMDIV requires 2 arguments")
	}
	inumber1, err := strconv.ParseComplex(str2cmplx(argsList.Front().Value.(formulaArg).Value()), 128)
	if err != nil {
		return newErrorFormulaArg(formulaErrorNUM, err.Error())
	}
	inumber2, err := strconv.ParseComplex(str2cmplx(argsList.Back().Value.(formulaArg).Value()), 128)
	if err != nil {
		return newErrorFormulaArg(formulaErrorNUM, err.Error())
	}
	num := inumber1 / inumber2
	if cmplx.IsInf(num) {
		return newErrorFormulaArg(formulaErrorNUM, formulaErrorNUM)
	}
	return newStringFormulaArg(cmplx2str(fmt.Sprint(num), "i"))
}

// IMEXP function returns the exponential of a supplied complex number. The
// syntax of the function is:
//
//    IMEXP(inumber)
//
func (fn *formulaFuncs) IMEXP(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "IMEXP requires 1 argument")
	}
	inumber, err := strconv.ParseComplex(str2cmplx(argsList.Front().Value.(formulaArg).Value()), 128)
	if err != nil {
		return newErrorFormulaArg(formulaErrorNUM, err.Error())
	}
	return newStringFormulaArg(cmplx2str(fmt.Sprint(cmplx.Exp(inumber)), "i"))
}

// IMLN function returns the natural logarithm of a supplied complex number.
// The syntax of the function is:
//
//    IMLN(inumber)
//
func (fn *formulaFuncs) IMLN(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "IMLN requires 1 argument")
	}
	inumber, err := strconv.ParseComplex(str2cmplx(argsList.Front().Value.(formulaArg).Value()), 128)
	if err != nil {
		return newErrorFormulaArg(formulaErrorNUM, err.Error())
	}
	num := cmplx.Log(inumber)
	if cmplx.IsInf(num) {
		return newErrorFormulaArg(formulaErrorNUM, formulaErrorNUM)
	}
	return newStringFormulaArg(cmplx2str(fmt.Sprint(num), "i"))
}

// IMLOG10 function returns the common (base 10) logarithm of a supplied
// complex number. The syntax of the function is:
//
//    IMLOG10(inumber)
//
func (fn *formulaFuncs) IMLOG10(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "IMLOG10 requires 1 argument")
	}
	inumber, err := strconv.ParseComplex(str2cmplx(argsList.Front().Value.(formulaArg).Value()), 128)
	if err != nil {
		return newErrorFormulaArg(formulaErrorNUM, err.Error())
	}
	num := cmplx.Log10(inumber)
	if cmplx.IsInf(num) {
		return newErrorFormulaArg(formulaErrorNUM, formulaErrorNUM)
	}
	return newStringFormulaArg(cmplx2str(fmt.Sprint(num), "i"))
}

// IMLOG2 function calculates the base 2 logarithm of a supplied complex
// number. The syntax of the function is:
//
//    IMLOG2(inumber)
//
func (fn *formulaFuncs) IMLOG2(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "IMLOG2 requires 1 argument")
	}
	inumber, err := strconv.ParseComplex(str2cmplx(argsList.Front().Value.(formulaArg).Value()), 128)
	if err != nil {
		return newErrorFormulaArg(formulaErrorNUM, err.Error())
	}
	num := cmplx.Log(inumber)
	if cmplx.IsInf(num) {
		return newErrorFormulaArg(formulaErrorNUM, formulaErrorNUM)
	}
	return newStringFormulaArg(cmplx2str(fmt.Sprint(num/cmplx.Log(2)), "i"))
}

// IMPOWER function returns a supplied complex number, raised to a given
// power. The syntax of the function is:
//
//    IMPOWER(inumber,number)
//
func (fn *formulaFuncs) IMPOWER(argsList *list.List) formulaArg {
	if argsList.Len() != 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "IMPOWER requires 2 arguments")
	}
	inumber, err := strconv.ParseComplex(str2cmplx(argsList.Front().Value.(formulaArg).Value()), 128)
	if err != nil {
		return newErrorFormulaArg(formulaErrorNUM, err.Error())
	}
	number, err := strconv.ParseComplex(str2cmplx(argsList.Back().Value.(formulaArg).Value()), 128)
	if err != nil {
		return newErrorFormulaArg(formulaErrorNUM, err.Error())
	}
	if inumber == 0 && number == 0 {
		return newErrorFormulaArg(formulaErrorNUM, formulaErrorNUM)
	}
	num := cmplx.Pow(inumber, number)
	if cmplx.IsInf(num) {
		return newErrorFormulaArg(formulaErrorNUM, formulaErrorNUM)
	}
	return newStringFormulaArg(cmplx2str(fmt.Sprint(num), "i"))
}

// IMPRODUCT function calculates the product of two or more complex numbers.
// The syntax of the function is:
//
//    IMPRODUCT(number1,[number2],...)
//
func (fn *formulaFuncs) IMPRODUCT(argsList *list.List) formulaArg {
	product := complex128(1)
	for arg := argsList.Front(); arg != nil; arg = arg.Next() {
		token := arg.Value.(formulaArg)
		switch token.Type {
		case ArgString:
			if token.Value() == "" {
				continue
			}
			val, err := strconv.ParseComplex(str2cmplx(token.Value()), 128)
			if err != nil {
				return newErrorFormulaArg(formulaErrorNUM, err.Error())
			}
			product = product * val
		case ArgNumber:
			product = product * complex(token.Number, 0)
		case ArgMatrix:
			for _, row := range token.Matrix {
				for _, value := range row {
					if value.Value() == "" {
						continue
					}
					val, err := strconv.ParseComplex(str2cmplx(value.Value()), 128)
					if err != nil {
						return newErrorFormulaArg(formulaErrorNUM, err.Error())
					}
					product = product * val
				}
			}
		}
	}
	return newStringFormulaArg(cmplx2str(fmt.Sprint(product), "i"))
}

// IMREAL function returns the real coefficient of a supplied complex number.
// The syntax of the function is:
//
//    IMREAL(inumber)
//
func (fn *formulaFuncs) IMREAL(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "IMREAL requires 1 argument")
	}
	inumber, err := strconv.ParseComplex(str2cmplx(argsList.Front().Value.(formulaArg).Value()), 128)
	if err != nil {
		return newErrorFormulaArg(formulaErrorNUM, err.Error())
	}
	return newStringFormulaArg(cmplx2str(fmt.Sprint(real(inumber)), "i"))
}

// IMSEC function returns the secant of a supplied complex number. The syntax
// of the function is:
//
//    IMSEC(inumber)
//
func (fn *formulaFuncs) IMSEC(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "IMSEC requires 1 argument")
	}
	inumber, err := strconv.ParseComplex(str2cmplx(argsList.Front().Value.(formulaArg).Value()), 128)
	if err != nil {
		return newErrorFormulaArg(formulaErrorNUM, err.Error())
	}
	return newStringFormulaArg(cmplx2str(fmt.Sprint(1/cmplx.Cos(inumber)), "i"))
}

// IMSECH function returns the hyperbolic secant of a supplied complex number.
// The syntax of the function is:
//
//    IMSECH(inumber)
//
func (fn *formulaFuncs) IMSECH(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "IMSECH requires 1 argument")
	}
	inumber, err := strconv.ParseComplex(str2cmplx(argsList.Front().Value.(formulaArg).Value()), 128)
	if err != nil {
		return newErrorFormulaArg(formulaErrorNUM, err.Error())
	}
	return newStringFormulaArg(cmplx2str(fmt.Sprint(1/cmplx.Cosh(inumber)), "i"))
}

// IMSIN function returns the Sine of a supplied complex number. The syntax of
// the function is:
//
//    IMSIN(inumber)
//
func (fn *formulaFuncs) IMSIN(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "IMSIN requires 1 argument")
	}
	inumber, err := strconv.ParseComplex(str2cmplx(argsList.Front().Value.(formulaArg).Value()), 128)
	if err != nil {
		return newErrorFormulaArg(formulaErrorNUM, err.Error())
	}
	return newStringFormulaArg(cmplx2str(fmt.Sprint(cmplx.Sin(inumber)), "i"))
}

// IMSINH function returns the hyperbolic sine of a supplied complex number.
// The syntax of the function is:
//
//    IMSINH(inumber)
//
func (fn *formulaFuncs) IMSINH(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "IMSINH requires 1 argument")
	}
	inumber, err := strconv.ParseComplex(str2cmplx(argsList.Front().Value.(formulaArg).Value()), 128)
	if err != nil {
		return newErrorFormulaArg(formulaErrorNUM, err.Error())
	}
	return newStringFormulaArg(cmplx2str(fmt.Sprint(cmplx.Sinh(inumber)), "i"))
}

// IMSQRT function returns the square root of a supplied complex number. The
// syntax of the function is:
//
//    IMSQRT(inumber)
//
func (fn *formulaFuncs) IMSQRT(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "IMSQRT requires 1 argument")
	}
	inumber, err := strconv.ParseComplex(str2cmplx(argsList.Front().Value.(formulaArg).Value()), 128)
	if err != nil {
		return newErrorFormulaArg(formulaErrorNUM, err.Error())
	}
	return newStringFormulaArg(cmplx2str(fmt.Sprint(cmplx.Sqrt(inumber)), "i"))
}

// IMSUB function calculates the difference between two complex numbers
// (i.e. subtracts one complex number from another). The syntax of the
// function is:
//
//    IMSUB(inumber1,inumber2)
//
func (fn *formulaFuncs) IMSUB(argsList *list.List) formulaArg {
	if argsList.Len() != 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "IMSUB requires 2 arguments")
	}
	i1, err := strconv.ParseComplex(str2cmplx(argsList.Front().Value.(formulaArg).Value()), 128)
	if err != nil {
		return newErrorFormulaArg(formulaErrorNUM, err.Error())
	}
	i2, err := strconv.ParseComplex(str2cmplx(argsList.Back().Value.(formulaArg).Value()), 128)
	if err != nil {
		return newErrorFormulaArg(formulaErrorNUM, err.Error())
	}
	return newStringFormulaArg(cmplx2str(fmt.Sprint(i1-i2), "i"))
}

// IMSUM function calculates the sum of two or more complex numbers. The
// syntax of the function is:
//
//    IMSUM(inumber1,inumber2,...)
//
func (fn *formulaFuncs) IMSUM(argsList *list.List) formulaArg {
	if argsList.Len() < 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "IMSUM requires at least 1 argument")
	}
	var result complex128
	for arg := argsList.Front(); arg != nil; arg = arg.Next() {
		token := arg.Value.(formulaArg)
		num, err := strconv.ParseComplex(str2cmplx(token.Value()), 128)
		if err != nil {
			return newErrorFormulaArg(formulaErrorNUM, err.Error())
		}
		result += num
	}
	return newStringFormulaArg(cmplx2str(fmt.Sprint(result), "i"))
}

// IMTAN function returns the tangent of a supplied complex number. The syntax
// of the function is:
//
//    IMTAN(inumber)
//
func (fn *formulaFuncs) IMTAN(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "IMTAN requires 1 argument")
	}
	inumber, err := strconv.ParseComplex(str2cmplx(argsList.Front().Value.(formulaArg).Value()), 128)
	if err != nil {
		return newErrorFormulaArg(formulaErrorNUM, err.Error())
	}
	return newStringFormulaArg(cmplx2str(fmt.Sprint(cmplx.Tan(inumber)), "i"))
}

// OCT2BIN function converts an Octal (Base 8) number into a Binary (Base 2)
// number. The syntax of the function is:
//
//    OCT2BIN(number,[places])
//
func (fn *formulaFuncs) OCT2BIN(argsList *list.List) formulaArg {
	if argsList.Len() < 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "OCT2BIN requires at least 1 argument")
	}
	if argsList.Len() > 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "OCT2BIN allows at most 2 arguments")
	}
	token := argsList.Front().Value.(formulaArg)
	number := token.ToNumber()
	if number.Type != ArgNumber {
		return newErrorFormulaArg(formulaErrorVALUE, number.Error)
	}
	decimal, newList := fn.oct2dec(token.Value()), list.New()
	newList.PushBack(decimal)
	if argsList.Len() == 2 {
		newList.PushBack(argsList.Back().Value.(formulaArg))
	}
	return fn.dec2x("OCT2BIN", newList)
}

// OCT2DEC function converts an Octal (a base-8 number) into a decimal number.
// The syntax of the function is:
//
//    OCT2DEC(number)
//
func (fn *formulaFuncs) OCT2DEC(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "OCT2DEC requires 1 numeric argument")
	}
	token := argsList.Front().Value.(formulaArg)
	number := token.ToNumber()
	if number.Type != ArgNumber {
		return newErrorFormulaArg(formulaErrorVALUE, number.Error)
	}
	return fn.oct2dec(token.Value())
}

// OCT2HEX function converts an Octal (Base 8) number into a Hexadecimal
// (Base 16) number. The syntax of the function is:
//
//    OCT2HEX(number,[places])
//
func (fn *formulaFuncs) OCT2HEX(argsList *list.List) formulaArg {
	if argsList.Len() < 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "OCT2HEX requires at least 1 argument")
	}
	if argsList.Len() > 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "OCT2HEX allows at most 2 arguments")
	}
	token := argsList.Front().Value.(formulaArg)
	number := token.ToNumber()
	if number.Type != ArgNumber {
		return newErrorFormulaArg(formulaErrorVALUE, number.Error)
	}
	decimal, newList := fn.oct2dec(token.Value()), list.New()
	newList.PushBack(decimal)
	if argsList.Len() == 2 {
		newList.PushBack(argsList.Back().Value.(formulaArg))
	}
	return fn.dec2x("OCT2HEX", newList)
}

// oct2dec is an implementation of the formula function OCT2DEC.
func (fn *formulaFuncs) oct2dec(number string) formulaArg {
	decimal, length := 0.0, len(number)
	for i := length; i > 0; i-- {
		num, _ := strconv.Atoi(string(number[length-i]))
		if i == 10 && string(number[length-i]) == "7" {
			decimal += math.Pow(-8.0, float64(i-1))
			continue
		}
		decimal += float64(num) * math.Pow(8.0, float64(i-1))
	}
	return newNumberFormulaArg(decimal)
}

// Math and Trigonometric Functions

// ABS function returns the absolute value of any supplied number. The syntax
// of the function is:
//
//    ABS(number)
//
func (fn *formulaFuncs) ABS(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "ABS requires 1 numeric argument")
	}
	arg := argsList.Front().Value.(formulaArg).ToNumber()
	if arg.Type == ArgError {
		return arg
	}
	return newNumberFormulaArg(math.Abs(arg.Number))
}

// ACOS function calculates the arccosine (i.e. the inverse cosine) of a given
// number, and returns an angle, in radians, between 0 and π. The syntax of
// the function is:
//
//    ACOS(number)
//
func (fn *formulaFuncs) ACOS(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "ACOS requires 1 numeric argument")
	}
	arg := argsList.Front().Value.(formulaArg).ToNumber()
	if arg.Type == ArgError {
		return arg
	}
	return newNumberFormulaArg(math.Acos(arg.Number))
}

// ACOSH function calculates the inverse hyperbolic cosine of a supplied number.
// of the function is:
//
//    ACOSH(number)
//
func (fn *formulaFuncs) ACOSH(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "ACOSH requires 1 numeric argument")
	}
	arg := argsList.Front().Value.(formulaArg).ToNumber()
	if arg.Type == ArgError {
		return arg
	}
	return newNumberFormulaArg(math.Acosh(arg.Number))
}

// ACOT function calculates the arccotangent (i.e. the inverse cotangent) of a
// given number, and returns an angle, in radians, between 0 and π. The syntax
// of the function is:
//
//    ACOT(number)
//
func (fn *formulaFuncs) ACOT(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "ACOT requires 1 numeric argument")
	}
	arg := argsList.Front().Value.(formulaArg).ToNumber()
	if arg.Type == ArgError {
		return arg
	}
	return newNumberFormulaArg(math.Pi/2 - math.Atan(arg.Number))
}

// ACOTH function calculates the hyperbolic arccotangent (coth) of a supplied
// value. The syntax of the function is:
//
//    ACOTH(number)
//
func (fn *formulaFuncs) ACOTH(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "ACOTH requires 1 numeric argument")
	}
	arg := argsList.Front().Value.(formulaArg).ToNumber()
	if arg.Type == ArgError {
		return arg
	}
	return newNumberFormulaArg(math.Atanh(1 / arg.Number))
}

// ARABIC function converts a Roman numeral into an Arabic numeral. The syntax
// of the function is:
//
//    ARABIC(text)
//
func (fn *formulaFuncs) ARABIC(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "ARABIC requires 1 numeric argument")
	}
	text := argsList.Front().Value.(formulaArg).Value()
	if len(text) > 255 {
		return newErrorFormulaArg(formulaErrorVALUE, formulaErrorVALUE)
	}
	text = strings.ToUpper(text)
	number, actualStart, index, isNegative := 0, 0, len(text)-1, false
	startIndex, subtractNumber, currentPartValue, currentCharValue, prevCharValue := 0, 0, 0, 0, -1
	for index >= 0 && text[index] == ' ' {
		index--
	}
	for actualStart <= index && text[actualStart] == ' ' {
		actualStart++
	}
	if actualStart <= index && text[actualStart] == '-' {
		isNegative = true
		actualStart++
	}
	charMap := map[rune]int{'I': 1, 'V': 5, 'X': 10, 'L': 50, 'C': 100, 'D': 500, 'M': 1000}
	for index >= actualStart {
		startIndex = index
		startChar := text[startIndex]
		index--
		for index >= actualStart && (text[index]|' ') == startChar {
			index--
		}
		currentCharValue = charMap[rune(startChar)]
		currentPartValue = (startIndex - index) * currentCharValue
		if currentCharValue >= prevCharValue {
			number += currentPartValue - subtractNumber
			prevCharValue = currentCharValue
			subtractNumber = 0
			continue
		}
		subtractNumber += currentPartValue
	}
	if subtractNumber != 0 {
		number -= subtractNumber
	}
	if isNegative {
		number = -number
	}
	return newNumberFormulaArg(float64(number))
}

// ASIN function calculates the arcsine (i.e. the inverse sine) of a given
// number, and returns an angle, in radians, between -π/2 and π/2. The syntax
// of the function is:
//
//    ASIN(number)
//
func (fn *formulaFuncs) ASIN(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "ASIN requires 1 numeric argument")
	}
	arg := argsList.Front().Value.(formulaArg).ToNumber()
	if arg.Type == ArgError {
		return arg
	}
	return newNumberFormulaArg(math.Asin(arg.Number))
}

// ASINH function calculates the inverse hyperbolic sine of a supplied number.
// The syntax of the function is:
//
//    ASINH(number)
//
func (fn *formulaFuncs) ASINH(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "ASINH requires 1 numeric argument")
	}
	arg := argsList.Front().Value.(formulaArg).ToNumber()
	if arg.Type == ArgError {
		return arg
	}
	return newNumberFormulaArg(math.Asinh(arg.Number))
}

// ATAN function calculates the arctangent (i.e. the inverse tangent) of a
// given number, and returns an angle, in radians, between -π/2 and +π/2. The
// syntax of the function is:
//
//    ATAN(number)
//
func (fn *formulaFuncs) ATAN(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "ATAN requires 1 numeric argument")
	}
	arg := argsList.Front().Value.(formulaArg).ToNumber()
	if arg.Type == ArgError {
		return arg
	}
	return newNumberFormulaArg(math.Atan(arg.Number))
}

// ATANH function calculates the inverse hyperbolic tangent of a supplied
// number. The syntax of the function is:
//
//    ATANH(number)
//
func (fn *formulaFuncs) ATANH(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "ATANH requires 1 numeric argument")
	}
	arg := argsList.Front().Value.(formulaArg).ToNumber()
	if arg.Type == ArgError {
		return arg
	}
	return newNumberFormulaArg(math.Atanh(arg.Number))
}

// ATAN2 function calculates the arctangent (i.e. the inverse tangent) of a
// given set of x and y coordinates, and returns an angle, in radians, between
// -π/2 and +π/2. The syntax of the function is:
//
//    ATAN2(x_num,y_num)
//
func (fn *formulaFuncs) ATAN2(argsList *list.List) formulaArg {
	if argsList.Len() != 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "ATAN2 requires 2 numeric arguments")
	}
	x := argsList.Back().Value.(formulaArg).ToNumber()
	if x.Type == ArgError {
		return x
	}
	y := argsList.Front().Value.(formulaArg).ToNumber()
	if y.Type == ArgError {
		return y
	}
	return newNumberFormulaArg(math.Atan2(x.Number, y.Number))
}

// BASE function converts a number into a supplied base (radix), and returns a
// text representation of the calculated value. The syntax of the function is:
//
//    BASE(number,radix,[min_length])
//
func (fn *formulaFuncs) BASE(argsList *list.List) formulaArg {
	if argsList.Len() < 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "BASE requires at least 2 arguments")
	}
	if argsList.Len() > 3 {
		return newErrorFormulaArg(formulaErrorVALUE, "BASE allows at most 3 arguments")
	}
	var minLength int
	var err error
	number := argsList.Front().Value.(formulaArg).ToNumber()
	if number.Type == ArgError {
		return number
	}
	radix := argsList.Front().Next().Value.(formulaArg).ToNumber()
	if radix.Type == ArgError {
		return radix
	}
	if int(radix.Number) < 2 || int(radix.Number) > 36 {
		return newErrorFormulaArg(formulaErrorVALUE, "radix must be an integer >= 2 and <= 36")
	}
	if argsList.Len() > 2 {
		if minLength, err = strconv.Atoi(argsList.Back().Value.(formulaArg).String); err != nil {
			return newErrorFormulaArg(formulaErrorVALUE, err.Error())
		}
	}
	result := strconv.FormatInt(int64(number.Number), int(radix.Number))
	if len(result) < minLength {
		result = strings.Repeat("0", minLength-len(result)) + result
	}
	return newStringFormulaArg(strings.ToUpper(result))
}

// CEILING function rounds a supplied number away from zero, to the nearest
// multiple of a given number. The syntax of the function is:
//
//    CEILING(number,significance)
//
func (fn *formulaFuncs) CEILING(argsList *list.List) formulaArg {
	if argsList.Len() == 0 {
		return newErrorFormulaArg(formulaErrorVALUE, "CEILING requires at least 1 argument")
	}
	if argsList.Len() > 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "CEILING allows at most 2 arguments")
	}
	number, significance, res := 0.0, 1.0, 0.0
	n := argsList.Front().Value.(formulaArg).ToNumber()
	if n.Type == ArgError {
		return n
	}
	number = n.Number
	if number < 0 {
		significance = -1
	}
	if argsList.Len() > 1 {
		s := argsList.Back().Value.(formulaArg).ToNumber()
		if s.Type == ArgError {
			return s
		}
		significance = s.Number
	}
	if significance < 0 && number > 0 {
		return newErrorFormulaArg(formulaErrorVALUE, "negative sig to CEILING invalid")
	}
	if argsList.Len() == 1 {
		return newNumberFormulaArg(math.Ceil(number))
	}
	number, res = math.Modf(number / significance)
	if res > 0 {
		number++
	}
	return newNumberFormulaArg(number * significance)
}

// CEILINGdotMATH function rounds a supplied number up to a supplied multiple
// of significance. The syntax of the function is:
//
//    CEILING.MATH(number,[significance],[mode])
//
func (fn *formulaFuncs) CEILINGdotMATH(argsList *list.List) formulaArg {
	if argsList.Len() == 0 {
		return newErrorFormulaArg(formulaErrorVALUE, "CEILING.MATH requires at least 1 argument")
	}
	if argsList.Len() > 3 {
		return newErrorFormulaArg(formulaErrorVALUE, "CEILING.MATH allows at most 3 arguments")
	}
	number, significance, mode := 0.0, 1.0, 1.0
	n := argsList.Front().Value.(formulaArg).ToNumber()
	if n.Type == ArgError {
		return n
	}
	number = n.Number
	if number < 0 {
		significance = -1
	}
	if argsList.Len() > 1 {
		s := argsList.Front().Next().Value.(formulaArg).ToNumber()
		if s.Type == ArgError {
			return s
		}
		significance = s.Number
	}
	if argsList.Len() == 1 {
		return newNumberFormulaArg(math.Ceil(number))
	}
	if argsList.Len() > 2 {
		m := argsList.Back().Value.(formulaArg).ToNumber()
		if m.Type == ArgError {
			return m
		}
		mode = m.Number
	}
	val, res := math.Modf(number / significance)
	if res != 0 {
		if number > 0 {
			val++
		} else if mode < 0 {
			val--
		}
	}
	return newNumberFormulaArg(val * significance)
}

// CEILINGdotPRECISE function rounds a supplied number up (regardless of the
// number's sign), to the nearest multiple of a given number. The syntax of
// the function is:
//
//    CEILING.PRECISE(number,[significance])
//
func (fn *formulaFuncs) CEILINGdotPRECISE(argsList *list.List) formulaArg {
	if argsList.Len() == 0 {
		return newErrorFormulaArg(formulaErrorVALUE, "CEILING.PRECISE requires at least 1 argument")
	}
	if argsList.Len() > 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "CEILING.PRECISE allows at most 2 arguments")
	}
	number, significance := 0.0, 1.0
	n := argsList.Front().Value.(formulaArg).ToNumber()
	if n.Type == ArgError {
		return n
	}
	number = n.Number
	if number < 0 {
		significance = -1
	}
	if argsList.Len() == 1 {
		return newNumberFormulaArg(math.Ceil(number))
	}
	if argsList.Len() > 1 {
		s := argsList.Back().Value.(formulaArg).ToNumber()
		if s.Type == ArgError {
			return s
		}
		significance = s.Number
		significance = math.Abs(significance)
		if significance == 0 {
			return newNumberFormulaArg(significance)
		}
	}
	val, res := math.Modf(number / significance)
	if res != 0 {
		if number > 0 {
			val++
		}
	}
	return newNumberFormulaArg(val * significance)
}

// COMBIN function calculates the number of combinations (in any order) of a
// given number objects from a set. The syntax of the function is:
//
//    COMBIN(number,number_chosen)
//
func (fn *formulaFuncs) COMBIN(argsList *list.List) formulaArg {
	if argsList.Len() != 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "COMBIN requires 2 argument")
	}
	number, chosen, val := 0.0, 0.0, 1.0
	n := argsList.Front().Value.(formulaArg).ToNumber()
	if n.Type == ArgError {
		return n
	}
	number = n.Number
	c := argsList.Back().Value.(formulaArg).ToNumber()
	if c.Type == ArgError {
		return c
	}
	chosen = c.Number
	number, chosen = math.Trunc(number), math.Trunc(chosen)
	if chosen > number {
		return newErrorFormulaArg(formulaErrorVALUE, "COMBIN requires number >= number_chosen")
	}
	if chosen == number || chosen == 0 {
		return newNumberFormulaArg(1)
	}
	for c := float64(1); c <= chosen; c++ {
		val *= (number + 1 - c) / c
	}
	return newNumberFormulaArg(math.Ceil(val))
}

// COMBINA function calculates the number of combinations, with repetitions,
// of a given number objects from a set. The syntax of the function is:
//
//    COMBINA(number,number_chosen)
//
func (fn *formulaFuncs) COMBINA(argsList *list.List) formulaArg {
	if argsList.Len() != 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "COMBINA requires 2 argument")
	}
	var number, chosen float64
	n := argsList.Front().Value.(formulaArg).ToNumber()
	if n.Type == ArgError {
		return n
	}
	number = n.Number
	c := argsList.Back().Value.(formulaArg).ToNumber()
	if c.Type == ArgError {
		return c
	}
	chosen = c.Number
	number, chosen = math.Trunc(number), math.Trunc(chosen)
	if number < chosen {
		return newErrorFormulaArg(formulaErrorVALUE, "COMBINA requires number > number_chosen")
	}
	if number == 0 {
		return newNumberFormulaArg(number)
	}
	args := list.New()
	args.PushBack(formulaArg{
		String: fmt.Sprintf("%g", number+chosen-1),
		Type:   ArgString,
	})
	args.PushBack(formulaArg{
		String: fmt.Sprintf("%g", number-1),
		Type:   ArgString,
	})
	return fn.COMBIN(args)
}

// COS function calculates the cosine of a given angle. The syntax of the
// function is:
//
//    COS(number)
//
func (fn *formulaFuncs) COS(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "COS requires 1 numeric argument")
	}
	val := argsList.Front().Value.(formulaArg).ToNumber()
	if val.Type == ArgError {
		return val
	}
	return newNumberFormulaArg(math.Cos(val.Number))
}

// COSH function calculates the hyperbolic cosine (cosh) of a supplied number.
// The syntax of the function is:
//
//    COSH(number)
//
func (fn *formulaFuncs) COSH(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "COSH requires 1 numeric argument")
	}
	val := argsList.Front().Value.(formulaArg).ToNumber()
	if val.Type == ArgError {
		return val
	}
	return newNumberFormulaArg(math.Cosh(val.Number))
}

// COT function calculates the cotangent of a given angle. The syntax of the
// function is:
//
//    COT(number)
//
func (fn *formulaFuncs) COT(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "COT requires 1 numeric argument")
	}
	val := argsList.Front().Value.(formulaArg).ToNumber()
	if val.Type == ArgError {
		return val
	}
	if val.Number == 0 {
		return newErrorFormulaArg(formulaErrorDIV, formulaErrorDIV)
	}
	return newNumberFormulaArg(1 / math.Tan(val.Number))
}

// COTH function calculates the hyperbolic cotangent (coth) of a supplied
// angle. The syntax of the function is:
//
//    COTH(number)
//
func (fn *formulaFuncs) COTH(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "COTH requires 1 numeric argument")
	}
	val := argsList.Front().Value.(formulaArg).ToNumber()
	if val.Type == ArgError {
		return val
	}
	if val.Number == 0 {
		return newErrorFormulaArg(formulaErrorDIV, formulaErrorDIV)
	}
	return newNumberFormulaArg((math.Exp(val.Number) + math.Exp(-val.Number)) / (math.Exp(val.Number) - math.Exp(-val.Number)))
}

// CSC function calculates the cosecant of a given angle. The syntax of the
// function is:
//
//    CSC(number)
//
func (fn *formulaFuncs) CSC(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "CSC requires 1 numeric argument")
	}
	val := argsList.Front().Value.(formulaArg).ToNumber()
	if val.Type == ArgError {
		return val
	}
	if val.Number == 0 {
		return newErrorFormulaArg(formulaErrorDIV, formulaErrorDIV)
	}
	return newNumberFormulaArg(1 / math.Sin(val.Number))
}

// CSCH function calculates the hyperbolic cosecant (csch) of a supplied
// angle. The syntax of the function is:
//
//    CSCH(number)
//
func (fn *formulaFuncs) CSCH(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "CSCH requires 1 numeric argument")
	}
	val := argsList.Front().Value.(formulaArg).ToNumber()
	if val.Type == ArgError {
		return val
	}
	if val.Number == 0 {
		return newErrorFormulaArg(formulaErrorDIV, formulaErrorDIV)
	}
	return newNumberFormulaArg(1 / math.Sinh(val.Number))
}

// DECIMAL function converts a text representation of a number in a specified
// base, into a decimal value. The syntax of the function is:
//
//    DECIMAL(text,radix)
//
func (fn *formulaFuncs) DECIMAL(argsList *list.List) formulaArg {
	if argsList.Len() != 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "DECIMAL requires 2 numeric arguments")
	}
	var text = argsList.Front().Value.(formulaArg).String
	var radix int
	var err error
	radix, err = strconv.Atoi(argsList.Back().Value.(formulaArg).String)
	if err != nil {
		return newErrorFormulaArg(formulaErrorVALUE, err.Error())
	}
	if len(text) > 2 && (strings.HasPrefix(text, "0x") || strings.HasPrefix(text, "0X")) {
		text = text[2:]
	}
	val, err := strconv.ParseInt(text, radix, 64)
	if err != nil {
		return newErrorFormulaArg(formulaErrorVALUE, err.Error())
	}
	return newNumberFormulaArg(float64(val))
}

// DEGREES function converts radians into degrees. The syntax of the function
// is:
//
//    DEGREES(angle)
//
func (fn *formulaFuncs) DEGREES(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "DEGREES requires 1 numeric argument")
	}
	val := argsList.Front().Value.(formulaArg).ToNumber()
	if val.Type == ArgError {
		return val
	}
	if val.Number == 0 {
		return newErrorFormulaArg(formulaErrorDIV, formulaErrorDIV)
	}
	return newNumberFormulaArg(180.0 / math.Pi * val.Number)
}

// EVEN function rounds a supplied number away from zero (i.e. rounds a
// positive number up and a negative number down), to the next even number.
// The syntax of the function is:
//
//    EVEN(number)
//
func (fn *formulaFuncs) EVEN(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "EVEN requires 1 numeric argument")
	}
	number := argsList.Front().Value.(formulaArg).ToNumber()
	if number.Type == ArgError {
		return number
	}
	sign := math.Signbit(number.Number)
	m, frac := math.Modf(number.Number / 2)
	val := m * 2
	if frac != 0 {
		if !sign {
			val += 2
		} else {
			val -= 2
		}
	}
	return newNumberFormulaArg(val)
}

// EXP function calculates the value of the mathematical constant e, raised to
// the power of a given number. The syntax of the function is:
//
//    EXP(number)
//
func (fn *formulaFuncs) EXP(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "EXP requires 1 numeric argument")
	}
	number := argsList.Front().Value.(formulaArg).ToNumber()
	if number.Type == ArgError {
		return number
	}
	return newStringFormulaArg(strings.ToUpper(fmt.Sprintf("%g", math.Exp(number.Number))))
}

// fact returns the factorial of a supplied number.
func fact(number float64) float64 {
	val := float64(1)
	for i := float64(2); i <= number; i++ {
		val *= i
	}
	return val
}

// FACT function returns the factorial of a supplied number. The syntax of the
// function is:
//
//    FACT(number)
//
func (fn *formulaFuncs) FACT(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "FACT requires 1 numeric argument")
	}
	number := argsList.Front().Value.(formulaArg).ToNumber()
	if number.Type == ArgError {
		return number
	}
	if number.Number < 0 {
		return newErrorFormulaArg(formulaErrorNUM, formulaErrorNUM)
	}
	return newNumberFormulaArg(fact(number.Number))
}

// FACTDOUBLE function returns the double factorial of a supplied number. The
// syntax of the function is:
//
//    FACTDOUBLE(number)
//
func (fn *formulaFuncs) FACTDOUBLE(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "FACTDOUBLE requires 1 numeric argument")
	}
	val := 1.0
	number := argsList.Front().Value.(formulaArg).ToNumber()
	if number.Type == ArgError {
		return number
	}
	if number.Number < 0 {
		return newErrorFormulaArg(formulaErrorNUM, formulaErrorNUM)
	}
	for i := math.Trunc(number.Number); i > 1; i -= 2 {
		val *= i
	}
	return newStringFormulaArg(strings.ToUpper(fmt.Sprintf("%g", val)))
}

// FLOOR function rounds a supplied number towards zero to the nearest
// multiple of a specified significance. The syntax of the function is:
//
//    FLOOR(number,significance)
//
func (fn *formulaFuncs) FLOOR(argsList *list.List) formulaArg {
	if argsList.Len() != 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "FLOOR requires 2 numeric arguments")
	}
	number := argsList.Front().Value.(formulaArg).ToNumber()
	if number.Type == ArgError {
		return number
	}
	significance := argsList.Back().Value.(formulaArg).ToNumber()
	if significance.Type == ArgError {
		return significance
	}
	if significance.Number < 0 && number.Number >= 0 {
		return newErrorFormulaArg(formulaErrorNUM, "invalid arguments to FLOOR")
	}
	val := number.Number
	val, res := math.Modf(val / significance.Number)
	if res != 0 {
		if number.Number < 0 && res < 0 {
			val--
		}
	}
	return newStringFormulaArg(strings.ToUpper(fmt.Sprintf("%g", val*significance.Number)))
}

// FLOORdotMATH function rounds a supplied number down to a supplied multiple
// of significance. The syntax of the function is:
//
//    FLOOR.MATH(number,[significance],[mode])
//
func (fn *formulaFuncs) FLOORdotMATH(argsList *list.List) formulaArg {
	if argsList.Len() == 0 {
		return newErrorFormulaArg(formulaErrorVALUE, "FLOOR.MATH requires at least 1 argument")
	}
	if argsList.Len() > 3 {
		return newErrorFormulaArg(formulaErrorVALUE, "FLOOR.MATH allows at most 3 arguments")
	}
	significance, mode := 1.0, 1.0
	number := argsList.Front().Value.(formulaArg).ToNumber()
	if number.Type == ArgError {
		return number
	}
	if number.Number < 0 {
		significance = -1
	}
	if argsList.Len() > 1 {
		s := argsList.Front().Next().Value.(formulaArg).ToNumber()
		if s.Type == ArgError {
			return s
		}
		significance = s.Number
	}
	if argsList.Len() == 1 {
		return newNumberFormulaArg(math.Floor(number.Number))
	}
	if argsList.Len() > 2 {
		m := argsList.Back().Value.(formulaArg).ToNumber()
		if m.Type == ArgError {
			return m
		}
		mode = m.Number
	}
	val, res := math.Modf(number.Number / significance)
	if res != 0 && number.Number < 0 && mode > 0 {
		val--
	}
	return newNumberFormulaArg(val * significance)
}

// FLOORdotPRECISE function rounds a supplied number down to a supplied
// multiple of significance. The syntax of the function is:
//
//    FLOOR.PRECISE(number,[significance])
//
func (fn *formulaFuncs) FLOORdotPRECISE(argsList *list.List) formulaArg {
	if argsList.Len() == 0 {
		return newErrorFormulaArg(formulaErrorVALUE, "FLOOR.PRECISE requires at least 1 argument")
	}
	if argsList.Len() > 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "FLOOR.PRECISE allows at most 2 arguments")
	}
	var significance float64
	number := argsList.Front().Value.(formulaArg).ToNumber()
	if number.Type == ArgError {
		return number
	}
	if number.Number < 0 {
		significance = -1
	}
	if argsList.Len() == 1 {
		return newNumberFormulaArg(math.Floor(number.Number))
	}
	if argsList.Len() > 1 {
		s := argsList.Back().Value.(formulaArg).ToNumber()
		if s.Type == ArgError {
			return s
		}
		significance = s.Number
		significance = math.Abs(significance)
		if significance == 0 {
			return newNumberFormulaArg(significance)
		}
	}
	val, res := math.Modf(number.Number / significance)
	if res != 0 {
		if number.Number < 0 {
			val--
		}
	}
	return newNumberFormulaArg(val * significance)
}

// gcd returns the greatest common divisor of two supplied integers.
func gcd(x, y float64) float64 {
	x, y = math.Trunc(x), math.Trunc(y)
	if x == 0 {
		return y
	}
	if y == 0 {
		return x
	}
	for x != y {
		if x > y {
			x = x - y
		} else {
			y = y - x
		}
	}
	return x
}

// GCD function returns the greatest common divisor of two or more supplied
// integers. The syntax of the function is:
//
//    GCD(number1,[number2],...)
//
func (fn *formulaFuncs) GCD(argsList *list.List) formulaArg {
	if argsList.Len() == 0 {
		return newErrorFormulaArg(formulaErrorVALUE, "GCD requires at least 1 argument")
	}
	var (
		val  float64
		nums = []float64{}
	)
	for arg := argsList.Front(); arg != nil; arg = arg.Next() {
		token := arg.Value.(formulaArg)
		switch token.Type {
		case ArgString:
			num := token.ToNumber()
			if num.Type == ArgError {
				return num
			}
			val = num.Number
		case ArgNumber:
			val = token.Number
		}
		nums = append(nums, val)
	}
	if nums[0] < 0 {
		return newErrorFormulaArg(formulaErrorVALUE, "GCD only accepts positive arguments")
	}
	if len(nums) == 1 {
		return newNumberFormulaArg(nums[0])
	}
	cd := nums[0]
	for i := 1; i < len(nums); i++ {
		if nums[i] < 0 {
			return newErrorFormulaArg(formulaErrorVALUE, "GCD only accepts positive arguments")
		}
		cd = gcd(cd, nums[i])
	}
	return newNumberFormulaArg(cd)
}

// INT function truncates a supplied number down to the closest integer. The
// syntax of the function is:
//
//    INT(number)
//
func (fn *formulaFuncs) INT(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "INT requires 1 numeric argument")
	}
	number := argsList.Front().Value.(formulaArg).ToNumber()
	if number.Type == ArgError {
		return number
	}
	val, frac := math.Modf(number.Number)
	if frac < 0 {
		val--
	}
	return newNumberFormulaArg(val)
}

// ISOdotCEILING function rounds a supplied number up (regardless of the
// number's sign), to the nearest multiple of a supplied significance. The
// syntax of the function is:
//
//    ISO.CEILING(number,[significance])
//
func (fn *formulaFuncs) ISOdotCEILING(argsList *list.List) formulaArg {
	if argsList.Len() == 0 {
		return newErrorFormulaArg(formulaErrorVALUE, "ISO.CEILING requires at least 1 argument")
	}
	if argsList.Len() > 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "ISO.CEILING allows at most 2 arguments")
	}
	var significance float64
	number := argsList.Front().Value.(formulaArg).ToNumber()
	if number.Type == ArgError {
		return number
	}
	if number.Number < 0 {
		significance = -1
	}
	if argsList.Len() == 1 {
		return newNumberFormulaArg(math.Ceil(number.Number))
	}
	if argsList.Len() > 1 {
		s := argsList.Back().Value.(formulaArg).ToNumber()
		if s.Type == ArgError {
			return s
		}
		significance = s.Number
		significance = math.Abs(significance)
		if significance == 0 {
			return newNumberFormulaArg(significance)
		}
	}
	val, res := math.Modf(number.Number / significance)
	if res != 0 {
		if number.Number > 0 {
			val++
		}
	}
	return newNumberFormulaArg(val * significance)
}

// lcm returns the least common multiple of two supplied integers.
func lcm(a, b float64) float64 {
	a = math.Trunc(a)
	b = math.Trunc(b)
	if a == 0 && b == 0 {
		return 0
	}
	return a * b / gcd(a, b)
}

// LCM function returns the least common multiple of two or more supplied
// integers. The syntax of the function is:
//
//    LCM(number1,[number2],...)
//
func (fn *formulaFuncs) LCM(argsList *list.List) formulaArg {
	if argsList.Len() == 0 {
		return newErrorFormulaArg(formulaErrorVALUE, "LCM requires at least 1 argument")
	}
	var (
		val  float64
		nums = []float64{}
		err  error
	)
	for arg := argsList.Front(); arg != nil; arg = arg.Next() {
		token := arg.Value.(formulaArg)
		switch token.Type {
		case ArgString:
			if token.String == "" {
				continue
			}
			if val, err = strconv.ParseFloat(token.String, 64); err != nil {
				return newErrorFormulaArg(formulaErrorVALUE, err.Error())
			}
		case ArgNumber:
			val = token.Number
		}
		nums = append(nums, val)
	}
	if nums[0] < 0 {
		return newErrorFormulaArg(formulaErrorVALUE, "LCM only accepts positive arguments")
	}
	if len(nums) == 1 {
		return newNumberFormulaArg(nums[0])
	}
	cm := nums[0]
	for i := 1; i < len(nums); i++ {
		if nums[i] < 0 {
			return newErrorFormulaArg(formulaErrorVALUE, "LCM only accepts positive arguments")
		}
		cm = lcm(cm, nums[i])
	}
	return newNumberFormulaArg(cm)
}

// LN function calculates the natural logarithm of a given number. The syntax
// of the function is:
//
//    LN(number)
//
func (fn *formulaFuncs) LN(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "LN requires 1 numeric argument")
	}
	number := argsList.Front().Value.(formulaArg).ToNumber()
	if number.Type == ArgError {
		return number
	}
	return newNumberFormulaArg(math.Log(number.Number))
}

// LOG function calculates the logarithm of a given number, to a supplied
// base. The syntax of the function is:
//
//    LOG(number,[base])
//
func (fn *formulaFuncs) LOG(argsList *list.List) formulaArg {
	if argsList.Len() == 0 {
		return newErrorFormulaArg(formulaErrorVALUE, "LOG requires at least 1 argument")
	}
	if argsList.Len() > 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "LOG allows at most 2 arguments")
	}
	base := 10.0
	number := argsList.Front().Value.(formulaArg).ToNumber()
	if number.Type == ArgError {
		return number
	}
	if argsList.Len() > 1 {
		b := argsList.Back().Value.(formulaArg).ToNumber()
		if b.Type == ArgError {
			return b
		}
		base = b.Number
	}
	if number.Number == 0 {
		return newErrorFormulaArg(formulaErrorNUM, formulaErrorDIV)
	}
	if base == 0 {
		return newErrorFormulaArg(formulaErrorNUM, formulaErrorDIV)
	}
	if base == 1 {
		return newErrorFormulaArg(formulaErrorDIV, formulaErrorDIV)
	}
	return newNumberFormulaArg(math.Log(number.Number) / math.Log(base))
}

// LOG10 function calculates the base 10 logarithm of a given number. The
// syntax of the function is:
//
//    LOG10(number)
//
func (fn *formulaFuncs) LOG10(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "LOG10 requires 1 numeric argument")
	}
	number := argsList.Front().Value.(formulaArg).ToNumber()
	if number.Type == ArgError {
		return number
	}
	return newNumberFormulaArg(math.Log10(number.Number))
}

// minor function implement a minor of a matrix A is the determinant of some
// smaller square matrix.
func minor(sqMtx [][]float64, idx int) [][]float64 {
	ret := [][]float64{}
	for i := range sqMtx {
		if i == 0 {
			continue
		}
		row := []float64{}
		for j := range sqMtx {
			if j == idx {
				continue
			}
			row = append(row, sqMtx[i][j])
		}
		ret = append(ret, row)
	}
	return ret
}

// det determinant of the 2x2 matrix.
func det(sqMtx [][]float64) float64 {
	if len(sqMtx) == 2 {
		m00 := sqMtx[0][0]
		m01 := sqMtx[0][1]
		m10 := sqMtx[1][0]
		m11 := sqMtx[1][1]
		return m00*m11 - m10*m01
	}
	var res, sgn float64 = 0, 1
	for j := range sqMtx {
		res += sgn * sqMtx[0][j] * det(minor(sqMtx, j))
		sgn *= -1
	}
	return res
}

// MDETERM calculates the determinant of a square matrix. The
// syntax of the function is:
//
//    MDETERM(array)
//
func (fn *formulaFuncs) MDETERM(argsList *list.List) (result formulaArg) {
	var (
		num    float64
		numMtx = [][]float64{}
		err    error
		strMtx [][]formulaArg
	)
	if argsList.Len() < 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "MDETERM requires at least 1 argument")
	}
	strMtx = argsList.Front().Value.(formulaArg).Matrix
	var rows = len(strMtx)
	for _, row := range argsList.Front().Value.(formulaArg).Matrix {
		if len(row) != rows {
			return newErrorFormulaArg(formulaErrorVALUE, formulaErrorVALUE)
		}
		numRow := []float64{}
		for _, ele := range row {
			if num, err = strconv.ParseFloat(ele.String, 64); err != nil {
				return newErrorFormulaArg(formulaErrorVALUE, err.Error())
			}
			numRow = append(numRow, num)
		}
		numMtx = append(numMtx, numRow)
	}
	return newNumberFormulaArg(det(numMtx))
}

// MOD function returns the remainder of a division between two supplied
// numbers. The syntax of the function is:
//
//    MOD(number,divisor)
//
func (fn *formulaFuncs) MOD(argsList *list.List) formulaArg {
	if argsList.Len() != 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "MOD requires 2 numeric arguments")
	}
	number := argsList.Front().Value.(formulaArg).ToNumber()
	if number.Type == ArgError {
		return number
	}
	divisor := argsList.Back().Value.(formulaArg).ToNumber()
	if divisor.Type == ArgError {
		return divisor
	}
	if divisor.Number == 0 {
		return newErrorFormulaArg(formulaErrorDIV, "MOD divide by zero")
	}
	trunc, rem := math.Modf(number.Number / divisor.Number)
	if rem < 0 {
		trunc--
	}
	return newNumberFormulaArg(number.Number - divisor.Number*trunc)
}

// MROUND function rounds a supplied number up or down to the nearest multiple
// of a given number. The syntax of the function is:
//
//    MROUND(number,multiple)
//
func (fn *formulaFuncs) MROUND(argsList *list.List) formulaArg {
	if argsList.Len() != 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "MROUND requires 2 numeric arguments")
	}
	n := argsList.Front().Value.(formulaArg).ToNumber()
	if n.Type == ArgError {
		return n
	}
	multiple := argsList.Back().Value.(formulaArg).ToNumber()
	if multiple.Type == ArgError {
		return multiple
	}
	if multiple.Number == 0 {
		return newErrorFormulaArg(formulaErrorNUM, formulaErrorNUM)
	}
	if multiple.Number < 0 && n.Number > 0 ||
		multiple.Number > 0 && n.Number < 0 {
		return newErrorFormulaArg(formulaErrorNUM, formulaErrorNUM)
	}
	number, res := math.Modf(n.Number / multiple.Number)
	if math.Trunc(res+0.5) > 0 {
		number++
	}
	return newNumberFormulaArg(number * multiple.Number)
}

// MULTINOMIAL function calculates the ratio of the factorial of a sum of
// supplied values to the product of factorials of those values. The syntax of
// the function is:
//
//    MULTINOMIAL(number1,[number2],...)
//
func (fn *formulaFuncs) MULTINOMIAL(argsList *list.List) formulaArg {
	val, num, denom := 0.0, 0.0, 1.0
	var err error
	for arg := argsList.Front(); arg != nil; arg = arg.Next() {
		token := arg.Value.(formulaArg)
		switch token.Type {
		case ArgString:
			if token.String == "" {
				continue
			}
			if val, err = strconv.ParseFloat(token.String, 64); err != nil {
				return newErrorFormulaArg(formulaErrorVALUE, err.Error())
			}
		case ArgNumber:
			val = token.Number
		}
		num += val
		denom *= fact(val)
	}
	return newNumberFormulaArg(fact(num) / denom)
}

// MUNIT function returns the unit matrix for a specified dimension. The
// syntax of the function is:
//
//   MUNIT(dimension)
//
func (fn *formulaFuncs) MUNIT(argsList *list.List) (result formulaArg) {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "MUNIT requires 1 numeric argument")
	}
	dimension := argsList.Back().Value.(formulaArg).ToNumber()
	if dimension.Type == ArgError || dimension.Number < 0 {
		return newErrorFormulaArg(formulaErrorVALUE, dimension.Error)
	}
	matrix := make([][]formulaArg, 0, int(dimension.Number))
	for i := 0; i < int(dimension.Number); i++ {
		row := make([]formulaArg, int(dimension.Number))
		for j := 0; j < int(dimension.Number); j++ {
			if i == j {
				row[j] = newNumberFormulaArg(1.0)
			} else {
				row[j] = newNumberFormulaArg(0.0)
			}
		}
		matrix = append(matrix, row)
	}
	return newMatrixFormulaArg(matrix)
}

// ODD function ounds a supplied number away from zero (i.e. rounds a positive
// number up and a negative number down), to the next odd number. The syntax
// of the function is:
//
//   ODD(number)
//
func (fn *formulaFuncs) ODD(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "ODD requires 1 numeric argument")
	}
	number := argsList.Back().Value.(formulaArg).ToNumber()
	if number.Type == ArgError {
		return number
	}
	if number.Number == 0 {
		return newNumberFormulaArg(1)
	}
	sign := math.Signbit(number.Number)
	m, frac := math.Modf((number.Number - 1) / 2)
	val := m*2 + 1
	if frac != 0 {
		if !sign {
			val += 2
		} else {
			val -= 2
		}
	}
	return newNumberFormulaArg(val)
}

// PI function returns the value of the mathematical constant π (pi), accurate
// to 15 digits (14 decimal places). The syntax of the function is:
//
//   PI()
//
func (fn *formulaFuncs) PI(argsList *list.List) formulaArg {
	if argsList.Len() != 0 {
		return newErrorFormulaArg(formulaErrorVALUE, "PI accepts no arguments")
	}
	return newNumberFormulaArg(math.Pi)
}

// POWER function calculates a given number, raised to a supplied power.
// The syntax of the function is:
//
//    POWER(number,power)
//
func (fn *formulaFuncs) POWER(argsList *list.List) formulaArg {
	if argsList.Len() != 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "POWER requires 2 numeric arguments")
	}
	x := argsList.Front().Value.(formulaArg).ToNumber()
	if x.Type == ArgError {
		return x
	}
	y := argsList.Back().Value.(formulaArg).ToNumber()
	if y.Type == ArgError {
		return y
	}
	if x.Number == 0 && y.Number == 0 {
		return newErrorFormulaArg(formulaErrorNUM, formulaErrorNUM)
	}
	if x.Number == 0 && y.Number < 0 {
		return newErrorFormulaArg(formulaErrorDIV, formulaErrorDIV)
	}
	return newNumberFormulaArg(math.Pow(x.Number, y.Number))
}

// PRODUCT function returns the product (multiplication) of a supplied set of
// numerical values. The syntax of the function is:
//
//    PRODUCT(number1,[number2],...)
//
func (fn *formulaFuncs) PRODUCT(argsList *list.List) formulaArg {
	val, product := 0.0, 1.0
	var err error
	for arg := argsList.Front(); arg != nil; arg = arg.Next() {
		token := arg.Value.(formulaArg)
		switch token.Type {
		case ArgUnknown:
			continue
		case ArgString:
			if token.String == "" {
				continue
			}
			if val, err = strconv.ParseFloat(token.String, 64); err != nil {
				return newErrorFormulaArg(formulaErrorVALUE, err.Error())
			}
			product = product * val
		case ArgNumber:
			product = product * token.Number
		case ArgMatrix:
			for _, row := range token.Matrix {
				for _, value := range row {
					if value.String == "" {
						continue
					}
					if val, err = strconv.ParseFloat(value.String, 64); err != nil {
						return newErrorFormulaArg(formulaErrorVALUE, err.Error())
					}
					product = product * val
				}
			}
		}
	}
	return newNumberFormulaArg(product)
}

// QUOTIENT function returns the integer portion of a division between two
// supplied numbers. The syntax of the function is:
//
//   QUOTIENT(numerator,denominator)
//
func (fn *formulaFuncs) QUOTIENT(argsList *list.List) formulaArg {
	if argsList.Len() != 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "QUOTIENT requires 2 numeric arguments")
	}
	x := argsList.Front().Value.(formulaArg).ToNumber()
	if x.Type == ArgError {
		return x
	}
	y := argsList.Back().Value.(formulaArg).ToNumber()
	if y.Type == ArgError {
		return y
	}
	if y.Number == 0 {
		return newErrorFormulaArg(formulaErrorDIV, formulaErrorDIV)
	}
	return newNumberFormulaArg(math.Trunc(x.Number / y.Number))
}

// RADIANS function converts radians into degrees. The syntax of the function is:
//
//   RADIANS(angle)
//
func (fn *formulaFuncs) RADIANS(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "RADIANS requires 1 numeric argument")
	}
	angle := argsList.Front().Value.(formulaArg).ToNumber()
	if angle.Type == ArgError {
		return angle
	}
	return newNumberFormulaArg(math.Pi / 180.0 * angle.Number)
}

// RAND function generates a random real number between 0 and 1. The syntax of
// the function is:
//
//   RAND()
//
func (fn *formulaFuncs) RAND(argsList *list.List) formulaArg {
	if argsList.Len() != 0 {
		return newErrorFormulaArg(formulaErrorVALUE, "RAND accepts no arguments")
	}
	return newNumberFormulaArg(rand.New(rand.NewSource(time.Now().UnixNano())).Float64())
}

// RANDBETWEEN function generates a random integer between two supplied
// integers. The syntax of the function is:
//
//   RANDBETWEEN(bottom,top)
//
func (fn *formulaFuncs) RANDBETWEEN(argsList *list.List) formulaArg {
	if argsList.Len() != 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "RANDBETWEEN requires 2 numeric arguments")
	}
	bottom := argsList.Front().Value.(formulaArg).ToNumber()
	if bottom.Type == ArgError {
		return bottom
	}
	top := argsList.Back().Value.(formulaArg).ToNumber()
	if top.Type == ArgError {
		return top
	}
	if top.Number < bottom.Number {
		return newErrorFormulaArg(formulaErrorNUM, formulaErrorNUM)
	}
	num := rand.New(rand.NewSource(time.Now().UnixNano())).Int63n(int64(top.Number - bottom.Number + 1))
	return newNumberFormulaArg(float64(num + int64(bottom.Number)))
}

// romanNumerals defined a numeral system that originated in ancient Rome and
// remained the usual way of writing numbers throughout Europe well into the
// Late Middle Ages.
type romanNumerals struct {
	n float64
	s string
}

var romanTable = [][]romanNumerals{
	{
		{1000, "M"}, {900, "CM"}, {500, "D"}, {400, "CD"}, {100, "C"}, {90, "XC"},
		{50, "L"}, {40, "XL"}, {10, "X"}, {9, "IX"}, {5, "V"}, {4, "IV"}, {1, "I"},
	},
	{
		{1000, "M"}, {950, "LM"}, {900, "CM"}, {500, "D"}, {450, "LD"}, {400, "CD"},
		{100, "C"}, {95, "VC"}, {90, "XC"}, {50, "L"}, {45, "VL"}, {40, "XL"},
		{10, "X"}, {9, "IX"}, {5, "V"}, {4, "IV"}, {1, "I"},
	},
	{
		{1000, "M"}, {990, "XM"}, {950, "LM"}, {900, "CM"}, {500, "D"}, {490, "XD"},
		{450, "LD"}, {400, "CD"}, {100, "C"}, {99, "IC"}, {90, "XC"}, {50, "L"},
		{45, "VL"}, {40, "XL"}, {10, "X"}, {9, "IX"}, {5, "V"}, {4, "IV"}, {1, "I"},
	},
	{
		{1000, "M"}, {995, "VM"}, {990, "XM"}, {950, "LM"}, {900, "CM"}, {500, "D"},
		{495, "VD"}, {490, "XD"}, {450, "LD"}, {400, "CD"}, {100, "C"}, {99, "IC"},
		{90, "XC"}, {50, "L"}, {45, "VL"}, {40, "XL"}, {10, "X"}, {9, "IX"},
		{5, "V"}, {4, "IV"}, {1, "I"},
	},
	{
		{1000, "M"}, {999, "IM"}, {995, "VM"}, {990, "XM"}, {950, "LM"}, {900, "CM"},
		{500, "D"}, {499, "ID"}, {495, "VD"}, {490, "XD"}, {450, "LD"}, {400, "CD"},
		{100, "C"}, {99, "IC"}, {90, "XC"}, {50, "L"}, {45, "VL"}, {40, "XL"},
		{10, "X"}, {9, "IX"}, {5, "V"}, {4, "IV"}, {1, "I"},
	},
}

// ROMAN function converts an arabic number to Roman. I.e. for a supplied
// integer, the function returns a text string depicting the roman numeral
// form of the number. The syntax of the function is:
//
//   ROMAN(number,[form])
//
func (fn *formulaFuncs) ROMAN(argsList *list.List) formulaArg {
	if argsList.Len() == 0 {
		return newErrorFormulaArg(formulaErrorVALUE, "ROMAN requires at least 1 argument")
	}
	if argsList.Len() > 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "ROMAN allows at most 2 arguments")
	}
	var form int
	number := argsList.Front().Value.(formulaArg).ToNumber()
	if number.Type == ArgError {
		return number
	}
	if argsList.Len() > 1 {
		f := argsList.Back().Value.(formulaArg).ToNumber()
		if f.Type == ArgError {
			return f
		}
		form = int(f.Number)
		if form < 0 {
			form = 0
		} else if form > 4 {
			form = 4
		}
	}
	decimalTable := romanTable[0]
	switch form {
	case 1:
		decimalTable = romanTable[1]
	case 2:
		decimalTable = romanTable[2]
	case 3:
		decimalTable = romanTable[3]
	case 4:
		decimalTable = romanTable[4]
	}
	val := math.Trunc(number.Number)
	buf := bytes.Buffer{}
	for _, r := range decimalTable {
		for val >= r.n {
			buf.WriteString(r.s)
			val -= r.n
		}
	}
	return newStringFormulaArg(buf.String())
}

type roundMode byte

const (
	closest roundMode = iota
	down
	up
)

// round rounds a supplied number up or down.
func (fn *formulaFuncs) round(number, digits float64, mode roundMode) float64 {
	var significance float64
	if digits > 0 {
		significance = math.Pow(1/10.0, digits)
	} else {
		significance = math.Pow(10.0, -digits)
	}
	val, res := math.Modf(number / significance)
	switch mode {
	case closest:
		const eps = 0.499999999
		if res >= eps {
			val++
		} else if res <= -eps {
			val--
		}
	case down:
	case up:
		if res > 0 {
			val++
		} else if res < 0 {
			val--
		}
	}
	return val * significance
}

// ROUND function rounds a supplied number up or down, to a specified number
// of decimal places. The syntax of the function is:
//
//   ROUND(number,num_digits)
//
func (fn *formulaFuncs) ROUND(argsList *list.List) formulaArg {
	if argsList.Len() != 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "ROUND requires 2 numeric arguments")
	}
	number := argsList.Front().Value.(formulaArg).ToNumber()
	if number.Type == ArgError {
		return number
	}
	digits := argsList.Back().Value.(formulaArg).ToNumber()
	if digits.Type == ArgError {
		return digits
	}
	return newNumberFormulaArg(fn.round(number.Number, digits.Number, closest))
}

// ROUNDDOWN function rounds a supplied number down towards zero, to a
// specified number of decimal places. The syntax of the function is:
//
//   ROUNDDOWN(number,num_digits)
//
func (fn *formulaFuncs) ROUNDDOWN(argsList *list.List) formulaArg {
	if argsList.Len() != 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "ROUNDDOWN requires 2 numeric arguments")
	}
	number := argsList.Front().Value.(formulaArg).ToNumber()
	if number.Type == ArgError {
		return number
	}
	digits := argsList.Back().Value.(formulaArg).ToNumber()
	if digits.Type == ArgError {
		return digits
	}
	return newNumberFormulaArg(fn.round(number.Number, digits.Number, down))
}

// ROUNDUP function rounds a supplied number up, away from zero, to a
// specified number of decimal places. The syntax of the function is:
//
//   ROUNDUP(number,num_digits)
//
func (fn *formulaFuncs) ROUNDUP(argsList *list.List) formulaArg {
	if argsList.Len() != 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "ROUNDUP requires 2 numeric arguments")
	}
	number := argsList.Front().Value.(formulaArg).ToNumber()
	if number.Type == ArgError {
		return number
	}
	digits := argsList.Back().Value.(formulaArg).ToNumber()
	if digits.Type == ArgError {
		return digits
	}
	return newNumberFormulaArg(fn.round(number.Number, digits.Number, up))
}

// SEC function calculates the secant of a given angle. The syntax of the
// function is:
//
//    SEC(number)
//
func (fn *formulaFuncs) SEC(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "SEC requires 1 numeric argument")
	}
	number := argsList.Front().Value.(formulaArg).ToNumber()
	if number.Type == ArgError {
		return number
	}
	return newNumberFormulaArg(math.Cos(number.Number))
}

// SECH function calculates the hyperbolic secant (sech) of a supplied angle.
// The syntax of the function is:
//
//    SECH(number)
//
func (fn *formulaFuncs) SECH(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "SECH requires 1 numeric argument")
	}
	number := argsList.Front().Value.(formulaArg).ToNumber()
	if number.Type == ArgError {
		return number
	}
	return newNumberFormulaArg(1 / math.Cosh(number.Number))
}

// SIGN function returns the arithmetic sign (+1, -1 or 0) of a supplied
// number. I.e. if the number is positive, the Sign function returns +1, if
// the number is negative, the function returns -1 and if the number is 0
// (zero), the function returns 0. The syntax of the function is:
//
//   SIGN(number)
//
func (fn *formulaFuncs) SIGN(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "SIGN requires 1 numeric argument")
	}
	val := argsList.Front().Value.(formulaArg).ToNumber()
	if val.Type == ArgError {
		return val
	}
	if val.Number < 0 {
		return newNumberFormulaArg(-1)
	}
	if val.Number > 0 {
		return newNumberFormulaArg(1)
	}
	return newNumberFormulaArg(0)
}

// SIN function calculates the sine of a given angle. The syntax of the
// function is:
//
//    SIN(number)
//
func (fn *formulaFuncs) SIN(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "SIN requires 1 numeric argument")
	}
	number := argsList.Front().Value.(formulaArg).ToNumber()
	if number.Type == ArgError {
		return number
	}
	return newNumberFormulaArg(math.Sin(number.Number))
}

// SINH function calculates the hyperbolic sine (sinh) of a supplied number.
// The syntax of the function is:
//
//    SINH(number)
//
func (fn *formulaFuncs) SINH(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "SINH requires 1 numeric argument")
	}
	number := argsList.Front().Value.(formulaArg).ToNumber()
	if number.Type == ArgError {
		return number
	}
	return newNumberFormulaArg(math.Sinh(number.Number))
}

// SQRT function calculates the positive square root of a supplied number. The
// syntax of the function is:
//
//    SQRT(number)
//
func (fn *formulaFuncs) SQRT(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "SQRT requires 1 numeric argument")
	}
	value := argsList.Front().Value.(formulaArg).ToNumber()
	if value.Type == ArgError {
		return value
	}
	if value.Number < 0 {
		return newErrorFormulaArg(formulaErrorNUM, formulaErrorNUM)
	}
	return newNumberFormulaArg(math.Sqrt(value.Number))
}

// SQRTPI function returns the square root of a supplied number multiplied by
// the mathematical constant, π. The syntax of the function is:
//
//    SQRTPI(number)
//
func (fn *formulaFuncs) SQRTPI(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "SQRTPI requires 1 numeric argument")
	}
	number := argsList.Front().Value.(formulaArg).ToNumber()
	if number.Type == ArgError {
		return number
	}
	return newNumberFormulaArg(math.Sqrt(number.Number * math.Pi))
}

// STDEV function calculates the sample standard deviation of a supplied set
// of values. The syntax of the function is:
//
//    STDEV(number1,[number2],...)
//
func (fn *formulaFuncs) STDEV(argsList *list.List) formulaArg {
	if argsList.Len() < 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "STDEV requires at least 1 argument")
	}
	return fn.stdev(false, argsList)
}

// STDEVdotS function calculates the sample standard deviation of a supplied
// set of values. The syntax of the function is:
//
//    STDEV.S(number1,[number2],...)
//
func (fn *formulaFuncs) STDEVdotS(argsList *list.List) formulaArg {
	if argsList.Len() < 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "STDEV.S requires at least 1 argument")
	}
	return fn.stdev(false, argsList)
}

// STDEVA function estimates standard deviation based on a sample. The
// standard deviation is a measure of how widely values are dispersed from
// the average value (the mean). The syntax of the function is:
//
//    STDEVA(number1,[number2],...)
//
func (fn *formulaFuncs) STDEVA(argsList *list.List) formulaArg {
	if argsList.Len() < 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "STDEVA requires at least 1 argument")
	}
	return fn.stdev(true, argsList)
}

// stdev is an implementation of the formula function STDEV and STDEVA.
func (fn *formulaFuncs) stdev(stdeva bool, argsList *list.List) formulaArg {
	pow := func(result, count float64, n, m formulaArg) (float64, float64) {
		if result == -1 {
			result = math.Pow((n.Number - m.Number), 2)
		} else {
			result += math.Pow((n.Number - m.Number), 2)
		}
		count++
		return result, count
	}
	count, result := -1.0, -1.0
	var mean formulaArg
	if stdeva {
		mean = fn.AVERAGEA(argsList)
	} else {
		mean = fn.AVERAGE(argsList)
	}
	for arg := argsList.Front(); arg != nil; arg = arg.Next() {
		token := arg.Value.(formulaArg)
		switch token.Type {
		case ArgString, ArgNumber:
			if !stdeva && (token.Value() == "TRUE" || token.Value() == "FALSE") {
				continue
			} else if stdeva && (token.Value() == "TRUE" || token.Value() == "FALSE") {
				num := token.ToBool()
				if num.Type == ArgNumber {
					result, count = pow(result, count, num, mean)
					continue
				}
			} else {
				num := token.ToNumber()
				if num.Type == ArgNumber {
					result, count = pow(result, count, num, mean)
				}
			}
		case ArgList, ArgMatrix:
			for _, row := range token.ToList() {
				if row.Type == ArgNumber || row.Type == ArgString {
					if !stdeva && (row.Value() == "TRUE" || row.Value() == "FALSE") {
						continue
					} else if stdeva && (row.Value() == "TRUE" || row.Value() == "FALSE") {
						num := row.ToBool()
						if num.Type == ArgNumber {
							result, count = pow(result, count, num, mean)
							continue
						}
					} else {
						num := row.ToNumber()
						if num.Type == ArgNumber {
							result, count = pow(result, count, num, mean)
						}
					}
				}
			}
		}
	}
	if count > 0 && result >= 0 {
		return newNumberFormulaArg(math.Sqrt(result / count))
	}
	return newErrorFormulaArg(formulaErrorDIV, formulaErrorDIV)
}

// POISSONdotDIST function calculates the Poisson Probability Mass Function or
// the Cumulative Poisson Probability Function for a supplied set of
// parameters. The syntax of the function is:
//
//    POISSON.DIST(x,mean,cumulative)
//
func (fn *formulaFuncs) POISSONdotDIST(argsList *list.List) formulaArg {
	if argsList.Len() != 3 {
		return newErrorFormulaArg(formulaErrorVALUE, "POISSON.DIST requires 3 arguments")
	}
	return fn.POISSON(argsList)
}

// POISSON function calculates the Poisson Probability Mass Function or the
// Cumulative Poisson Probability Function for a supplied set of parameters.
// The syntax of the function is:
//
//    POISSON(x,mean,cumulative)
//
func (fn *formulaFuncs) POISSON(argsList *list.List) formulaArg {
	if argsList.Len() != 3 {
		return newErrorFormulaArg(formulaErrorVALUE, "POISSON requires 3 arguments")
	}
	var x, mean, cumulative formulaArg
	if x = argsList.Front().Value.(formulaArg).ToNumber(); x.Type != ArgNumber {
		return x
	}
	if mean = argsList.Front().Next().Value.(formulaArg).ToNumber(); mean.Type != ArgNumber {
		return mean
	}
	if cumulative = argsList.Back().Value.(formulaArg).ToBool(); cumulative.Type == ArgError {
		return cumulative
	}
	if x.Number < 0 || mean.Number <= 0 {
		return newErrorFormulaArg(formulaErrorNA, formulaErrorNA)
	}
	if cumulative.Number == 1 {
		summer := 0.0
		floor := math.Floor(x.Number)
		for i := 0; i <= int(floor); i++ {
			summer += math.Pow(mean.Number, float64(i)) / fact(float64(i))
		}
		return newNumberFormulaArg(math.Exp(0-mean.Number) * summer)
	}
	return newNumberFormulaArg(math.Exp(0-mean.Number) * math.Pow(mean.Number, x.Number) / fact(x.Number))
}

// SUM function adds together a supplied set of numbers and returns the sum of
// these values. The syntax of the function is:
//
//    SUM(number1,[number2],...)
//
func (fn *formulaFuncs) SUM(argsList *list.List) formulaArg {
	var sum float64
	for arg := argsList.Front(); arg != nil; arg = arg.Next() {
		token := arg.Value.(formulaArg)
		switch token.Type {
		case ArgUnknown:
			continue
		case ArgString:
			if num := token.ToNumber(); num.Type == ArgNumber {
				sum += num.Number
			}
		case ArgNumber:
			sum += token.Number
		case ArgMatrix:
			for _, row := range token.Matrix {
				for _, value := range row {
					if num := value.ToNumber(); num.Type == ArgNumber {
						sum += num.Number
					}
				}
			}
		}
	}
	return newNumberFormulaArg(sum)
}

// SUMIF function finds the values in a supplied array, that satisfy a given
// criteria, and returns the sum of the corresponding values in a second
// supplied array. The syntax of the function is:
//
//    SUMIF(range,criteria,[sum_range])
//
func (fn *formulaFuncs) SUMIF(argsList *list.List) formulaArg {
	if argsList.Len() < 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "SUMIF requires at least 2 argument")
	}
	var criteria = formulaCriteriaParser(argsList.Front().Next().Value.(formulaArg).String)
	var rangeMtx = argsList.Front().Value.(formulaArg).Matrix
	var sumRange [][]formulaArg
	if argsList.Len() == 3 {
		sumRange = argsList.Back().Value.(formulaArg).Matrix
	}
	var sum, val float64
	var err error
	for rowIdx, row := range rangeMtx {
		for colIdx, col := range row {
			var ok bool
			fromVal := col.String
			if col.String == "" {
				continue
			}
			if ok, err = formulaCriteriaEval(fromVal, criteria); err != nil {
				return newErrorFormulaArg(formulaErrorVALUE, err.Error())
			}
			if ok {
				if argsList.Len() == 3 {
					if len(sumRange) <= rowIdx || len(sumRange[rowIdx]) <= colIdx {
						continue
					}
					fromVal = sumRange[rowIdx][colIdx].String
				}
				if val, err = strconv.ParseFloat(fromVal, 64); err != nil {
					return newErrorFormulaArg(formulaErrorVALUE, err.Error())
				}
				sum += val
			}
		}
	}
	return newNumberFormulaArg(sum)
}

// SUMSQ function returns the sum of squares of a supplied set of values. The
// syntax of the function is:
//
//    SUMSQ(number1,[number2],...)
//
func (fn *formulaFuncs) SUMSQ(argsList *list.List) formulaArg {
	var val, sq float64
	var err error
	for arg := argsList.Front(); arg != nil; arg = arg.Next() {
		token := arg.Value.(formulaArg)
		switch token.Type {
		case ArgString:
			if token.String == "" {
				continue
			}
			if val, err = strconv.ParseFloat(token.String, 64); err != nil {
				return newErrorFormulaArg(formulaErrorVALUE, err.Error())
			}
			sq += val * val
		case ArgNumber:
			sq += token.Number
		case ArgMatrix:
			for _, row := range token.Matrix {
				for _, value := range row {
					if value.String == "" {
						continue
					}
					if val, err = strconv.ParseFloat(value.String, 64); err != nil {
						return newErrorFormulaArg(formulaErrorVALUE, err.Error())
					}
					sq += val * val
				}
			}
		}
	}
	return newNumberFormulaArg(sq)
}

// TAN function calculates the tangent of a given angle. The syntax of the
// function is:
//
//    TAN(number)
//
func (fn *formulaFuncs) TAN(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "TAN requires 1 numeric argument")
	}
	number := argsList.Front().Value.(formulaArg).ToNumber()
	if number.Type == ArgError {
		return number
	}
	return newNumberFormulaArg(math.Tan(number.Number))
}

// TANH function calculates the hyperbolic tangent (tanh) of a supplied
// number. The syntax of the function is:
//
//    TANH(number)
//
func (fn *formulaFuncs) TANH(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "TANH requires 1 numeric argument")
	}
	number := argsList.Front().Value.(formulaArg).ToNumber()
	if number.Type == ArgError {
		return number
	}
	return newNumberFormulaArg(math.Tanh(number.Number))
}

// TRUNC function truncates a supplied number to a specified number of decimal
// places. The syntax of the function is:
//
//    TRUNC(number,[number_digits])
//
func (fn *formulaFuncs) TRUNC(argsList *list.List) formulaArg {
	if argsList.Len() == 0 {
		return newErrorFormulaArg(formulaErrorVALUE, "TRUNC requires at least 1 argument")
	}
	var digits, adjust, rtrim float64
	var err error
	number := argsList.Front().Value.(formulaArg).ToNumber()
	if number.Type == ArgError {
		return number
	}
	if argsList.Len() > 1 {
		d := argsList.Back().Value.(formulaArg).ToNumber()
		if d.Type == ArgError {
			return d
		}
		digits = d.Number
		digits = math.Floor(digits)
	}
	adjust = math.Pow(10, digits)
	x := int((math.Abs(number.Number) - math.Abs(float64(int(number.Number)))) * adjust)
	if x != 0 {
		if rtrim, err = strconv.ParseFloat(strings.TrimRight(strconv.Itoa(x), "0"), 64); err != nil {
			return newErrorFormulaArg(formulaErrorVALUE, err.Error())
		}
	}
	if (digits > 0) && (rtrim < adjust/10) {
		return newNumberFormulaArg(number.Number)
	}
	return newNumberFormulaArg(float64(int(number.Number*adjust)) / adjust)
}

// Statistical Functions

// AVERAGE function returns the arithmetic mean of a list of supplied numbers.
// The syntax of the function is:
//
//    AVERAGE(number1,[number2],...)
//
func (fn *formulaFuncs) AVERAGE(argsList *list.List) formulaArg {
	args := []formulaArg{}
	for arg := argsList.Front(); arg != nil; arg = arg.Next() {
		args = append(args, arg.Value.(formulaArg))
	}
	count, sum := fn.countSum(false, args)
	if count == 0 {
		return newErrorFormulaArg(formulaErrorDIV, "AVERAGE divide by zero")
	}
	return newNumberFormulaArg(sum / count)
}

// AVERAGEA function returns the arithmetic mean of a list of supplied numbers
// with text cell and zero values. The syntax of the function is:
//
//    AVERAGEA(number1,[number2],...)
//
func (fn *formulaFuncs) AVERAGEA(argsList *list.List) formulaArg {
	args := []formulaArg{}
	for arg := argsList.Front(); arg != nil; arg = arg.Next() {
		args = append(args, arg.Value.(formulaArg))
	}
	count, sum := fn.countSum(true, args)
	if count == 0 {
		return newErrorFormulaArg(formulaErrorDIV, "AVERAGEA divide by zero")
	}
	return newNumberFormulaArg(sum / count)
}

// countSum get count and sum for a formula arguments array.
func (fn *formulaFuncs) countSum(countText bool, args []formulaArg) (count, sum float64) {
	for _, arg := range args {
		switch arg.Type {
		case ArgNumber:
			if countText || !arg.Boolean {
				sum += arg.Number
				count++
			}
		case ArgString:
			if !countText && (arg.Value() == "TRUE" || arg.Value() == "FALSE") {
				continue
			} else if countText && (arg.Value() == "TRUE" || arg.Value() == "FALSE") {
				num := arg.ToBool()
				if num.Type == ArgNumber {
					count++
					sum += num.Number
					continue
				}
			}
			num := arg.ToNumber()
			if countText && num.Type == ArgError && arg.String != "" {
				count++
			}
			if num.Type == ArgNumber {
				sum += num.Number
				count++
			}
		case ArgList, ArgMatrix:
			cnt, summary := fn.countSum(countText, arg.ToList())
			sum += summary
			count += cnt
		}
	}
	return
}

// COUNT function returns the count of numeric values in a supplied set of
// cells or values. This count includes both numbers and dates. The syntax of
// the function is:
//
//    COUNT(value1,[value2],...)
//
func (fn *formulaFuncs) COUNT(argsList *list.List) formulaArg {
	var count int
	for token := argsList.Front(); token != nil; token = token.Next() {
		arg := token.Value.(formulaArg)
		switch arg.Type {
		case ArgString:
			if arg.ToNumber().Type != ArgError {
				count++
			}
		case ArgNumber:
			count++
		case ArgMatrix:
			for _, row := range arg.Matrix {
				for _, value := range row {
					if value.ToNumber().Type != ArgError {
						count++
					}
				}
			}
		}
	}
	return newNumberFormulaArg(float64(count))
}

// COUNTA function returns the number of non-blanks within a supplied set of
// cells or values. The syntax of the function is:
//
//    COUNTA(value1,[value2],...)
//
func (fn *formulaFuncs) COUNTA(argsList *list.List) formulaArg {
	var count int
	for token := argsList.Front(); token != nil; token = token.Next() {
		arg := token.Value.(formulaArg)
		switch arg.Type {
		case ArgString:
			if arg.String != "" {
				count++
			}
		case ArgNumber:
			count++
		case ArgMatrix:
			for _, row := range arg.ToList() {
				switch row.Type {
				case ArgString:
					if row.String != "" {
						count++
					}
				case ArgNumber:
					count++
				}
			}
		}
	}
	return newNumberFormulaArg(float64(count))
}

// COUNTBLANK function returns the number of blank cells in a supplied range.
// The syntax of the function is:
//
//    COUNTBLANK(range)
//
func (fn *formulaFuncs) COUNTBLANK(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "COUNTBLANK requires 1 argument")
	}
	var count int
	token := argsList.Front().Value.(formulaArg)
	switch token.Type {
	case ArgString:
		if token.String == "" {
			count++
		}
	case ArgList, ArgMatrix:
		for _, row := range token.ToList() {
			switch row.Type {
			case ArgString:
				if row.String == "" {
					count++
				}
			case ArgEmpty:
				count++
			}
		}
	case ArgEmpty:
		count++
	}
	return newNumberFormulaArg(float64(count))
}

// FISHER function calculates the Fisher Transformation for a supplied value.
// The syntax of the function is:
//
//    FISHER(x)
//
func (fn *formulaFuncs) FISHER(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "FISHER requires 1 numeric argument")
	}
	token := argsList.Front().Value.(formulaArg)
	switch token.Type {
	case ArgString:
		arg := token.ToNumber()
		if arg.Type == ArgNumber {
			if arg.Number <= -1 || arg.Number >= 1 {
				return newErrorFormulaArg(formulaErrorNA, formulaErrorNA)
			}
			return newNumberFormulaArg(0.5 * math.Log((1+arg.Number)/(1-arg.Number)))
		}
	case ArgNumber:
		if token.Number <= -1 || token.Number >= 1 {
			return newErrorFormulaArg(formulaErrorNA, formulaErrorNA)
		}
		return newNumberFormulaArg(0.5 * math.Log((1+token.Number)/(1-token.Number)))
	}
	return newErrorFormulaArg(formulaErrorVALUE, "FISHER requires 1 numeric argument")
}

// FISHERINV function calculates the inverse of the Fisher Transformation and
// returns a value between -1 and +1. The syntax of the function is:
//
//    FISHERINV(y)
//
func (fn *formulaFuncs) FISHERINV(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "FISHERINV requires 1 numeric argument")
	}
	token := argsList.Front().Value.(formulaArg)
	switch token.Type {
	case ArgString:
		arg := token.ToNumber()
		if arg.Type == ArgNumber {
			return newNumberFormulaArg((math.Exp(2*arg.Number) - 1) / (math.Exp(2*arg.Number) + 1))
		}
	case ArgNumber:
		return newNumberFormulaArg((math.Exp(2*token.Number) - 1) / (math.Exp(2*token.Number) + 1))
	}
	return newErrorFormulaArg(formulaErrorVALUE, "FISHERINV requires 1 numeric argument")
}

// GAMMA function returns the value of the Gamma Function, Γ(n), for a
// specified number, n. The syntax of the function is:
//
//    GAMMA(number)
//
func (fn *formulaFuncs) GAMMA(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "GAMMA requires 1 numeric argument")
	}
	token := argsList.Front().Value.(formulaArg)
	switch token.Type {
	case ArgString:
		arg := token.ToNumber()
		if arg.Type == ArgNumber {
			if arg.Number <= 0 {
				return newErrorFormulaArg(formulaErrorNA, formulaErrorNA)
			}
			return newNumberFormulaArg(math.Gamma(arg.Number))
		}
	case ArgNumber:
		if token.Number <= 0 {
			return newErrorFormulaArg(formulaErrorNA, formulaErrorNA)
		}
		return newNumberFormulaArg(math.Gamma(token.Number))
	}
	return newErrorFormulaArg(formulaErrorVALUE, "GAMMA requires 1 numeric argument")
}

// GAMMALN function returns the natural logarithm of the Gamma Function, Γ
// (n). The syntax of the function is:
//
//    GAMMALN(x)
//
func (fn *formulaFuncs) GAMMALN(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "GAMMALN requires 1 numeric argument")
	}
	token := argsList.Front().Value.(formulaArg)
	switch token.Type {
	case ArgString:
		arg := token.ToNumber()
		if arg.Type == ArgNumber {
			if arg.Number <= 0 {
				return newErrorFormulaArg(formulaErrorNA, formulaErrorNA)
			}
			return newNumberFormulaArg(math.Log(math.Gamma(arg.Number)))
		}
	case ArgNumber:
		if token.Number <= 0 {
			return newErrorFormulaArg(formulaErrorNA, formulaErrorNA)
		}
		return newNumberFormulaArg(math.Log(math.Gamma(token.Number)))
	}
	return newErrorFormulaArg(formulaErrorVALUE, "GAMMALN requires 1 numeric argument")
}

// HARMEAN function calculates the harmonic mean of a supplied set of values.
// The syntax of the function is:
//
//    HARMEAN(number1,[number2],...)
//
func (fn *formulaFuncs) HARMEAN(argsList *list.List) formulaArg {
	if argsList.Len() < 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "HARMEAN requires at least 1 argument")
	}
	if min := fn.MIN(argsList); min.Number < 0 {
		return newErrorFormulaArg(formulaErrorNA, formulaErrorNA)
	}
	number, val, cnt := 0.0, 0.0, 0.0
	for token := argsList.Front(); token != nil; token = token.Next() {
		arg := token.Value.(formulaArg)
		switch arg.Type {
		case ArgString:
			num := arg.ToNumber()
			if num.Type != ArgNumber {
				continue
			}
			number = num.Number
		case ArgNumber:
			number = arg.Number
		}
		if number <= 0 {
			return newErrorFormulaArg(formulaErrorNA, formulaErrorNA)
		}
		val += (1 / number)
		cnt++
	}
	return newNumberFormulaArg(1 / (val / cnt))
}

// KURT function calculates the kurtosis of a supplied set of values. The
// syntax of the function is:
//
//    KURT(number1,[number2],...)
//
func (fn *formulaFuncs) KURT(argsList *list.List) formulaArg {
	if argsList.Len() < 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "KURT requires at least 1 argument")
	}
	mean, stdev := fn.AVERAGE(argsList), fn.STDEV(argsList)
	if stdev.Number > 0 {
		count, summer := 0.0, 0.0
		for arg := argsList.Front(); arg != nil; arg = arg.Next() {
			token := arg.Value.(formulaArg)
			switch token.Type {
			case ArgString, ArgNumber:
				num := token.ToNumber()
				if num.Type == ArgError {
					continue
				}
				summer += math.Pow((num.Number-mean.Number)/stdev.Number, 4)
				count++
			case ArgList, ArgMatrix:
				for _, row := range token.ToList() {
					if row.Type == ArgNumber || row.Type == ArgString {
						num := row.ToNumber()
						if num.Type == ArgError {
							continue
						}
						summer += math.Pow((num.Number-mean.Number)/stdev.Number, 4)
						count++
					}
				}
			}
		}
		if count > 3 {
			return newNumberFormulaArg(summer*(count*(count+1)/((count-1)*(count-2)*(count-3))) - (3 * math.Pow(count-1, 2) / ((count - 2) * (count - 3))))
		}
	}
	return newErrorFormulaArg(formulaErrorDIV, formulaErrorDIV)
}

// NORMdotDIST function calculates the Normal Probability Density Function or
// the Cumulative Normal Distribution. Function for a supplied set of
// parameters. The syntax of the function is:
//
//    NORM.DIST(x,mean,standard_dev,cumulative)
//
func (fn *formulaFuncs) NORMdotDIST(argsList *list.List) formulaArg {
	if argsList.Len() != 4 {
		return newErrorFormulaArg(formulaErrorVALUE, "NORM.DIST requires 4 arguments")
	}
	return fn.NORMDIST(argsList)
}

// NORMDIST function calculates the Normal Probability Density Function or the
// Cumulative Normal Distribution. Function for a supplied set of parameters.
// The syntax of the function is:
//
//    NORMDIST(x,mean,standard_dev,cumulative)
//
func (fn *formulaFuncs) NORMDIST(argsList *list.List) formulaArg {
	if argsList.Len() != 4 {
		return newErrorFormulaArg(formulaErrorVALUE, "NORMDIST requires 4 arguments")
	}
	var x, mean, stdDev, cumulative formulaArg
	if x = argsList.Front().Value.(formulaArg).ToNumber(); x.Type != ArgNumber {
		return x
	}
	if mean = argsList.Front().Next().Value.(formulaArg).ToNumber(); mean.Type != ArgNumber {
		return mean
	}
	if stdDev = argsList.Back().Prev().Value.(formulaArg).ToNumber(); stdDev.Type != ArgNumber {
		return stdDev
	}
	if cumulative = argsList.Back().Value.(formulaArg).ToBool(); cumulative.Type == ArgError {
		return cumulative
	}
	if stdDev.Number < 0 {
		return newErrorFormulaArg(formulaErrorNA, formulaErrorNA)
	}
	if cumulative.Number == 1 {
		return newNumberFormulaArg(0.5 * (1 + math.Erf((x.Number-mean.Number)/(stdDev.Number*math.Sqrt(2)))))
	}
	return newNumberFormulaArg((1 / (math.Sqrt(2*math.Pi) * stdDev.Number)) * math.Exp(0-(math.Pow(x.Number-mean.Number, 2)/(2*(stdDev.Number*stdDev.Number)))))
}

// NORMdotINV function calculates the inverse of the Cumulative Normal
// Distribution Function for a supplied value of x, and a supplied
// distribution mean & standard deviation. The syntax of the function is:
//
//    NORM.INV(probability,mean,standard_dev)
//
func (fn *formulaFuncs) NORMdotINV(argsList *list.List) formulaArg {
	if argsList.Len() != 3 {
		return newErrorFormulaArg(formulaErrorVALUE, "NORM.INV requires 3 arguments")
	}
	return fn.NORMINV(argsList)
}

// NORMINV function calculates the inverse of the Cumulative Normal
// Distribution Function for a supplied value of x, and a supplied
// distribution mean & standard deviation. The syntax of the function is:
//
//    NORMINV(probability,mean,standard_dev)
//
func (fn *formulaFuncs) NORMINV(argsList *list.List) formulaArg {
	if argsList.Len() != 3 {
		return newErrorFormulaArg(formulaErrorVALUE, "NORMINV requires 3 arguments")
	}
	var prob, mean, stdDev formulaArg
	if prob = argsList.Front().Value.(formulaArg).ToNumber(); prob.Type != ArgNumber {
		return prob
	}
	if mean = argsList.Front().Next().Value.(formulaArg).ToNumber(); mean.Type != ArgNumber {
		return mean
	}
	if stdDev = argsList.Back().Value.(formulaArg).ToNumber(); stdDev.Type != ArgNumber {
		return stdDev
	}
	if prob.Number < 0 || prob.Number > 1 {
		return newErrorFormulaArg(formulaErrorNA, formulaErrorNA)
	}
	if stdDev.Number < 0 {
		return newErrorFormulaArg(formulaErrorNA, formulaErrorNA)
	}
	inv, err := norminv(prob.Number)
	if err != nil {
		return newErrorFormulaArg(err.Error(), err.Error())
	}
	return newNumberFormulaArg(inv*stdDev.Number + mean.Number)
}

// NORMdotSdotDIST function calculates the Standard Normal Cumulative
// Distribution Function for a supplied value. The syntax of the function
// is:
//
//    NORM.S.DIST(z)
//
func (fn *formulaFuncs) NORMdotSdotDIST(argsList *list.List) formulaArg {
	if argsList.Len() != 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "NORM.S.DIST requires 2 numeric arguments")
	}
	args := list.New().Init()
	args.PushBack(argsList.Front().Value.(formulaArg))
	args.PushBack(formulaArg{Type: ArgNumber, Number: 0})
	args.PushBack(formulaArg{Type: ArgNumber, Number: 1})
	args.PushBack(argsList.Back().Value.(formulaArg))
	return fn.NORMDIST(args)
}

// NORMSDIST function calculates the Standard Normal Cumulative Distribution
// Function for a supplied value. The syntax of the function is:
//
//    NORMSDIST(z)
//
func (fn *formulaFuncs) NORMSDIST(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "NORMSDIST requires 1 numeric argument")
	}
	args := list.New().Init()
	args.PushBack(argsList.Front().Value.(formulaArg))
	args.PushBack(formulaArg{Type: ArgNumber, Number: 0})
	args.PushBack(formulaArg{Type: ArgNumber, Number: 1})
	args.PushBack(formulaArg{Type: ArgNumber, Number: 1, Boolean: true})
	return fn.NORMDIST(args)
}

// NORMSINV function calculates the inverse of the Standard Normal Cumulative
// Distribution Function for a supplied probability value. The syntax of the
// function is:
//
//    NORMSINV(probability)
//
func (fn *formulaFuncs) NORMSINV(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "NORMSINV requires 1 numeric argument")
	}
	args := list.New().Init()
	args.PushBack(argsList.Front().Value.(formulaArg))
	args.PushBack(formulaArg{Type: ArgNumber, Number: 0})
	args.PushBack(formulaArg{Type: ArgNumber, Number: 1})
	return fn.NORMINV(args)
}

// NORMdotSdotINV function calculates the inverse of the Standard Normal
// Cumulative Distribution Function for a supplied probability value. The
// syntax of the function is:
//
//    NORM.S.INV(probability)
//
func (fn *formulaFuncs) NORMdotSdotINV(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "NORM.S.INV requires 1 numeric argument")
	}
	args := list.New().Init()
	args.PushBack(argsList.Front().Value.(formulaArg))
	args.PushBack(formulaArg{Type: ArgNumber, Number: 0})
	args.PushBack(formulaArg{Type: ArgNumber, Number: 1})
	return fn.NORMINV(args)
}

// norminv returns the inverse of the normal cumulative distribution for the
// specified value.
func norminv(p float64) (float64, error) {
	a := map[int]float64{
		1: -3.969683028665376e+01, 2: 2.209460984245205e+02, 3: -2.759285104469687e+02,
		4: 1.383577518672690e+02, 5: -3.066479806614716e+01, 6: 2.506628277459239e+00,
	}
	b := map[int]float64{
		1: -5.447609879822406e+01, 2: 1.615858368580409e+02, 3: -1.556989798598866e+02,
		4: 6.680131188771972e+01, 5: -1.328068155288572e+01,
	}
	c := map[int]float64{
		1: -7.784894002430293e-03, 2: -3.223964580411365e-01, 3: -2.400758277161838e+00,
		4: -2.549732539343734e+00, 5: 4.374664141464968e+00, 6: 2.938163982698783e+00,
	}
	d := map[int]float64{
		1: 7.784695709041462e-03, 2: 3.224671290700398e-01, 3: 2.445134137142996e+00,
		4: 3.754408661907416e+00,
	}
	pLow := 0.02425   // Use lower region approx. below this
	pHigh := 1 - pLow // Use upper region approx. above this
	if 0 < p && p < pLow {
		// Rational approximation for lower region.
		q := math.Sqrt(-2 * math.Log(p))
		return (((((c[1]*q+c[2])*q+c[3])*q+c[4])*q+c[5])*q + c[6]) /
			((((d[1]*q+d[2])*q+d[3])*q+d[4])*q + 1), nil
	} else if pLow <= p && p <= pHigh {
		// Rational approximation for central region.
		q := p - 0.5
		r := q * q
		return (((((a[1]*r+a[2])*r+a[3])*r+a[4])*r+a[5])*r + a[6]) * q /
			(((((b[1]*r+b[2])*r+b[3])*r+b[4])*r+b[5])*r + 1), nil
	} else if pHigh < p && p < 1 {
		// Rational approximation for upper region.
		q := math.Sqrt(-2 * math.Log(1-p))
		return -(((((c[1]*q+c[2])*q+c[3])*q+c[4])*q+c[5])*q + c[6]) /
			((((d[1]*q+d[2])*q+d[3])*q+d[4])*q + 1), nil
	}
	return 0, errors.New(formulaErrorNUM)
}

// kth is an implementation of the formula function LARGE and SMALL.
func (fn *formulaFuncs) kth(name string, argsList *list.List) formulaArg {
	if argsList.Len() != 2 {
		return newErrorFormulaArg(formulaErrorVALUE, fmt.Sprintf("%s requires 2 arguments", name))
	}
	array := argsList.Front().Value.(formulaArg).ToList()
	kArg := argsList.Back().Value.(formulaArg).ToNumber()
	if kArg.Type != ArgNumber {
		return kArg
	}
	k := int(kArg.Number)
	if k < 1 {
		return newErrorFormulaArg(formulaErrorNUM, "k should be > 0")
	}
	data := []float64{}
	for _, arg := range array {
		if numArg := arg.ToNumber(); numArg.Type == ArgNumber {
			data = append(data, numArg.Number)
		}
	}
	if len(data) < k {
		return newErrorFormulaArg(formulaErrorNUM, "k should be <= length of array")
	}
	sort.Float64s(data)
	if name == "LARGE" {
		return newNumberFormulaArg(data[len(data)-k])
	}
	return newNumberFormulaArg(data[k-1])
}

// LARGE function returns the k'th largest value from an array of numeric
// values. The syntax of the function is:
//
//    LARGE(array,k)
//
func (fn *formulaFuncs) LARGE(argsList *list.List) formulaArg {
	return fn.kth("LARGE", argsList)
}

// MAX function returns the largest value from a supplied set of numeric
// values. The syntax of the function is:
//
//    MAX(number1,[number2],...)
//
func (fn *formulaFuncs) MAX(argsList *list.List) formulaArg {
	if argsList.Len() == 0 {
		return newErrorFormulaArg(formulaErrorVALUE, "MAX requires at least 1 argument")
	}
	return fn.max(false, argsList)
}

// MAXA function returns the largest value from a supplied set of numeric
// values, while counting text and the logical value FALSE as the value 0 and
// counting the logical value TRUE as the value 1. The syntax of the function
// is:
//
//    MAXA(number1,[number2],...)
//
func (fn *formulaFuncs) MAXA(argsList *list.List) formulaArg {
	if argsList.Len() == 0 {
		return newErrorFormulaArg(formulaErrorVALUE, "MAXA requires at least 1 argument")
	}
	return fn.max(true, argsList)
}

// max is an implementation of the formula function MAX and MAXA.
func (fn *formulaFuncs) max(maxa bool, argsList *list.List) formulaArg {
	max := -math.MaxFloat64
	for token := argsList.Front(); token != nil; token = token.Next() {
		arg := token.Value.(formulaArg)
		switch arg.Type {
		case ArgString:
			if !maxa && (arg.Value() == "TRUE" || arg.Value() == "FALSE") {
				continue
			} else {
				num := arg.ToBool()
				if num.Type == ArgNumber && num.Number > max {
					max = num.Number
					continue
				}
			}
			num := arg.ToNumber()
			if num.Type != ArgError && num.Number > max {
				max = num.Number
			}
		case ArgNumber:
			if arg.Number > max {
				max = arg.Number
			}
		case ArgList, ArgMatrix:
			for _, row := range arg.ToList() {
				switch row.Type {
				case ArgString:
					if !maxa && (row.Value() == "TRUE" || row.Value() == "FALSE") {
						continue
					} else {
						num := row.ToBool()
						if num.Type == ArgNumber && num.Number > max {
							max = num.Number
							continue
						}
					}
					num := row.ToNumber()
					if num.Type != ArgError && num.Number > max {
						max = num.Number
					}
				case ArgNumber:
					if row.Number > max {
						max = row.Number
					}
				}
			}
		case ArgError:
			return arg
		}
	}
	if max == -math.MaxFloat64 {
		max = 0
	}
	return newNumberFormulaArg(max)
}

// MEDIAN function returns the statistical median (the middle value) of a list
// of supplied numbers. The syntax of the function is:
//
//    MEDIAN(number1,[number2],...)
//
func (fn *formulaFuncs) MEDIAN(argsList *list.List) formulaArg {
	if argsList.Len() == 0 {
		return newErrorFormulaArg(formulaErrorVALUE, "MEDIAN requires at least 1 argument")
	}
	var values = []float64{}
	var median, digits float64
	var err error
	for token := argsList.Front(); token != nil; token = token.Next() {
		arg := token.Value.(formulaArg)
		switch arg.Type {
		case ArgString:
			num := arg.ToNumber()
			if num.Type == ArgError {
				return newErrorFormulaArg(formulaErrorVALUE, num.Error)
			}
			values = append(values, num.Number)
		case ArgNumber:
			values = append(values, arg.Number)
		case ArgMatrix:
			for _, row := range arg.Matrix {
				for _, value := range row {
					if value.String == "" {
						continue
					}
					if digits, err = strconv.ParseFloat(value.String, 64); err != nil {
						return newErrorFormulaArg(formulaErrorVALUE, err.Error())
					}
					values = append(values, digits)
				}
			}
		}
	}
	sort.Float64s(values)
	if len(values)%2 == 0 {
		median = (values[len(values)/2-1] + values[len(values)/2]) / 2
	} else {
		median = values[len(values)/2]
	}
	return newNumberFormulaArg(median)
}

// MIN function returns the smallest value from a supplied set of numeric
// values. The syntax of the function is:
//
//    MIN(number1,[number2],...)
//
func (fn *formulaFuncs) MIN(argsList *list.List) formulaArg {
	if argsList.Len() == 0 {
		return newErrorFormulaArg(formulaErrorVALUE, "MIN requires at least 1 argument")
	}
	return fn.min(false, argsList)
}

// MINA function returns the smallest value from a supplied set of numeric
// values, while counting text and the logical value FALSE as the value 0 and
// counting the logical value TRUE as the value 1. The syntax of the function
// is:
//
//    MINA(number1,[number2],...)
//
func (fn *formulaFuncs) MINA(argsList *list.List) formulaArg {
	if argsList.Len() == 0 {
		return newErrorFormulaArg(formulaErrorVALUE, "MINA requires at least 1 argument")
	}
	return fn.min(true, argsList)
}

// min is an implementation of the formula function MIN and MINA.
func (fn *formulaFuncs) min(mina bool, argsList *list.List) formulaArg {
	min := math.MaxFloat64
	for token := argsList.Front(); token != nil; token = token.Next() {
		arg := token.Value.(formulaArg)
		switch arg.Type {
		case ArgString:
			if !mina && (arg.Value() == "TRUE" || arg.Value() == "FALSE") {
				continue
			} else {
				num := arg.ToBool()
				if num.Type == ArgNumber && num.Number < min {
					min = num.Number
					continue
				}
			}
			num := arg.ToNumber()
			if num.Type != ArgError && num.Number < min {
				min = num.Number
			}
		case ArgNumber:
			if arg.Number < min {
				min = arg.Number
			}
		case ArgList, ArgMatrix:
			for _, row := range arg.ToList() {
				switch row.Type {
				case ArgString:
					if !mina && (row.Value() == "TRUE" || row.Value() == "FALSE") {
						continue
					} else {
						num := row.ToBool()
						if num.Type == ArgNumber && num.Number < min {
							min = num.Number
							continue
						}
					}
					num := row.ToNumber()
					if num.Type != ArgError && num.Number < min {
						min = num.Number
					}
				case ArgNumber:
					if row.Number < min {
						min = row.Number
					}
				}
			}
		case ArgError:
			return arg
		}
	}
	if min == math.MaxFloat64 {
		min = 0
	}
	return newNumberFormulaArg(min)
}

// PERCENTILEdotINC function returns the k'th percentile (i.e. the value below
// which k% of the data values fall) for a supplied range of values and a
// supplied k. The syntax of the function is:
//
//    PERCENTILE.INC(array,k)
//
func (fn *formulaFuncs) PERCENTILEdotINC(argsList *list.List) formulaArg {
	if argsList.Len() != 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "PERCENTILE.INC requires 2 arguments")
	}
	return fn.PERCENTILE(argsList)
}

// PERCENTILE function returns the k'th percentile (i.e. the value below which
// k% of the data values fall) for a supplied range of values and a supplied
// k. The syntax of the function is:
//
//    PERCENTILE(array,k)
//
func (fn *formulaFuncs) PERCENTILE(argsList *list.List) formulaArg {
	if argsList.Len() != 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "PERCENTILE requires 2 arguments")
	}
	array := argsList.Front().Value.(formulaArg).ToList()
	k := argsList.Back().Value.(formulaArg).ToNumber()
	if k.Type != ArgNumber {
		return k
	}
	if k.Number < 0 || k.Number > 1 {
		return newErrorFormulaArg(formulaErrorNA, formulaErrorNA)
	}
	numbers := []float64{}
	for _, arg := range array {
		if arg.Type == ArgError {
			return arg
		}
		num := arg.ToNumber()
		if num.Type == ArgNumber {
			numbers = append(numbers, num.Number)
		}
	}
	cnt := len(numbers)
	sort.Float64s(numbers)
	idx := k.Number * (float64(cnt) - 1)
	base := math.Floor(idx)
	if idx == base {
		return newNumberFormulaArg(numbers[int(idx)])
	}
	next := base + 1
	proportion := idx - base
	return newNumberFormulaArg(numbers[int(base)] + ((numbers[int(next)] - numbers[int(base)]) * proportion))
}

// PERMUT function calculates the number of permutations of a specified number
// of objects from a set of objects. The syntax of the function is:
//
//    PERMUT(number,number_chosen)
//
func (fn *formulaFuncs) PERMUT(argsList *list.List) formulaArg {
	if argsList.Len() != 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "PERMUT requires 2 numeric arguments")
	}
	number := argsList.Front().Value.(formulaArg).ToNumber()
	chosen := argsList.Back().Value.(formulaArg).ToNumber()
	if number.Type != ArgNumber {
		return number
	}
	if chosen.Type != ArgNumber {
		return chosen
	}
	if number.Number < chosen.Number {
		return newErrorFormulaArg(formulaErrorNA, formulaErrorNA)
	}
	return newNumberFormulaArg(math.Round(fact(number.Number) / fact(number.Number-chosen.Number)))
}

// PERMUTATIONA function calculates the number of permutations, with
// repetitions, of a specified number of objects from a set. The syntax of
// the function is:
//
//    PERMUTATIONA(number,number_chosen)
//
func (fn *formulaFuncs) PERMUTATIONA(argsList *list.List) formulaArg {
	if argsList.Len() < 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "PERMUTATIONA requires 2 numeric arguments")
	}
	number := argsList.Front().Value.(formulaArg).ToNumber()
	chosen := argsList.Back().Value.(formulaArg).ToNumber()
	if number.Type != ArgNumber {
		return number
	}
	if chosen.Type != ArgNumber {
		return chosen
	}
	num, numChosen := math.Floor(number.Number), math.Floor(chosen.Number)
	if num < 0 || numChosen < 0 {
		return newErrorFormulaArg(formulaErrorNA, formulaErrorNA)
	}
	return newNumberFormulaArg(math.Pow(num, numChosen))
}

// QUARTILE function returns a requested quartile of a supplied range of
// values. The syntax of the function is:
//
//    QUARTILE(array,quart)
//
func (fn *formulaFuncs) QUARTILE(argsList *list.List) formulaArg {
	if argsList.Len() != 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "QUARTILE requires 2 arguments")
	}
	quart := argsList.Back().Value.(formulaArg).ToNumber()
	if quart.Type != ArgNumber {
		return quart
	}
	if quart.Number < 0 || quart.Number > 4 {
		return newErrorFormulaArg(formulaErrorNUM, formulaErrorNUM)
	}
	args := list.New().Init()
	args.PushBack(argsList.Front().Value.(formulaArg))
	args.PushBack(newNumberFormulaArg(quart.Number / 4))
	return fn.PERCENTILE(args)
}

// QUARTILEdotINC function returns a requested quartile of a supplied range of
// values. The syntax of the function is:
//
//    QUARTILE.INC(array,quart)
//
func (fn *formulaFuncs) QUARTILEdotINC(argsList *list.List) formulaArg {
	if argsList.Len() != 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "QUARTILE.INC requires 2 arguments")
	}
	return fn.QUARTILE(argsList)
}

// SKEW function calculates the skewness of the distribution of a supplied set
// of values. The syntax of the function is:
//
//    SKEW(number1,[number2],...)
//
func (fn *formulaFuncs) SKEW(argsList *list.List) formulaArg {
	if argsList.Len() < 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "SKEW requires at least 1 argument")
	}
	mean, stdDev, count, summer := fn.AVERAGE(argsList), fn.STDEV(argsList), 0.0, 0.0
	for arg := argsList.Front(); arg != nil; arg = arg.Next() {
		token := arg.Value.(formulaArg)
		switch token.Type {
		case ArgNumber, ArgString:
			num := token.ToNumber()
			if num.Type == ArgError {
				return num
			}
			summer += math.Pow((num.Number-mean.Number)/stdDev.Number, 3)
			count++
		case ArgList, ArgMatrix:
			for _, row := range token.ToList() {
				numArg := row.ToNumber()
				if numArg.Type != ArgNumber {
					continue
				}
				summer += math.Pow((numArg.Number-mean.Number)/stdDev.Number, 3)
				count++
			}
		}
	}
	if count > 2 {
		return newNumberFormulaArg(summer * (count / ((count - 1) * (count - 2))))
	}
	return newErrorFormulaArg(formulaErrorDIV, formulaErrorDIV)
}

// SMALL function returns the k'th smallest value from an array of numeric
// values. The syntax of the function is:
//
//    SMALL(array,k)
//
func (fn *formulaFuncs) SMALL(argsList *list.List) formulaArg {
	return fn.kth("SMALL", argsList)
}

// VARP function returns the Variance of a given set of values. The syntax of
// the function is:
//
//    VARP(number1,[number2],...)
//
func (fn *formulaFuncs) VARP(argsList *list.List) formulaArg {
	if argsList.Len() < 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "VARP requires at least 1 argument")
	}
	summerA, summerB, count := 0.0, 0.0, 0.0
	for arg := argsList.Front(); arg != nil; arg = arg.Next() {
		for _, token := range arg.Value.(formulaArg).ToList() {
			if num := token.ToNumber(); num.Type == ArgNumber {
				summerA += (num.Number * num.Number)
				summerB += num.Number
				count++
			}
		}
	}
	if count > 0 {
		summerA *= count
		summerB *= summerB
		return newNumberFormulaArg((summerA - summerB) / (count * count))
	}
	return newErrorFormulaArg(formulaErrorDIV, formulaErrorDIV)
}

// VARdotP function returns the Variance of a given set of values. The syntax
// of the function is:
//
//    VAR.P(number1,[number2],...)
//
func (fn *formulaFuncs) VARdotP(argsList *list.List) formulaArg {
	if argsList.Len() < 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "VAR.P requires at least 1 argument")
	}
	return fn.VARP(argsList)
}

// Information Functions

// ISBLANK function tests if a specified cell is blank (empty) and if so,
// returns TRUE; Otherwise the function returns FALSE. The syntax of the
// function is:
//
//    ISBLANK(value)
//
func (fn *formulaFuncs) ISBLANK(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "ISBLANK requires 1 argument")
	}
	token := argsList.Front().Value.(formulaArg)
	result := "FALSE"
	switch token.Type {
	case ArgUnknown:
		result = "TRUE"
	case ArgString:
		if token.String == "" {
			result = "TRUE"
		}
	}
	return newStringFormulaArg(result)
}

// ISERR function tests if an initial supplied expression (or value) returns
// any Excel Error, except the #N/A error. If so, the function returns the
// logical value TRUE; If the supplied value is not an error or is the #N/A
// error, the ISERR function returns FALSE. The syntax of the function is:
//
//    ISERR(value)
//
func (fn *formulaFuncs) ISERR(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "ISERR requires 1 argument")
	}
	token := argsList.Front().Value.(formulaArg)
	result := "FALSE"
	if token.Type == ArgError {
		for _, errType := range []string{
			formulaErrorDIV, formulaErrorNAME, formulaErrorNUM,
			formulaErrorVALUE, formulaErrorREF, formulaErrorNULL,
			formulaErrorSPILL, formulaErrorCALC, formulaErrorGETTINGDATA,
		} {
			if errType == token.String {
				result = "TRUE"
			}
		}
	}
	return newStringFormulaArg(result)
}

// ISERROR function tests if an initial supplied expression (or value) returns
// an Excel Error, and if so, returns the logical value TRUE; Otherwise the
// function returns FALSE. The syntax of the function is:
//
//    ISERROR(value)
//
func (fn *formulaFuncs) ISERROR(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "ISERROR requires 1 argument")
	}
	token := argsList.Front().Value.(formulaArg)
	result := "FALSE"
	if token.Type == ArgError {
		for _, errType := range []string{
			formulaErrorDIV, formulaErrorNAME, formulaErrorNA, formulaErrorNUM,
			formulaErrorVALUE, formulaErrorREF, formulaErrorNULL, formulaErrorSPILL,
			formulaErrorCALC, formulaErrorGETTINGDATA,
		} {
			if errType == token.String {
				result = "TRUE"
			}
		}
	}
	return newStringFormulaArg(result)
}

// ISEVEN function tests if a supplied number (or numeric expression)
// evaluates to an even number, and if so, returns TRUE; Otherwise, the
// function returns FALSE. The syntax of the function is:
//
//    ISEVEN(value)
//
func (fn *formulaFuncs) ISEVEN(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "ISEVEN requires 1 argument")
	}
	var (
		token   = argsList.Front().Value.(formulaArg)
		result  = "FALSE"
		numeric int
		err     error
	)
	if token.Type == ArgString {
		if numeric, err = strconv.Atoi(token.String); err != nil {
			return newErrorFormulaArg(formulaErrorVALUE, err.Error())
		}
		if numeric == numeric/2*2 {
			return newStringFormulaArg("TRUE")
		}
	}
	return newStringFormulaArg(result)
}

// ISNA function tests if an initial supplied expression (or value) returns
// the Excel #N/A Error, and if so, returns TRUE; Otherwise the function
// returns FALSE. The syntax of the function is:
//
//    ISNA(value)
//
func (fn *formulaFuncs) ISNA(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "ISNA requires 1 argument")
	}
	token := argsList.Front().Value.(formulaArg)
	result := "FALSE"
	if token.Type == ArgError && token.String == formulaErrorNA {
		result = "TRUE"
	}
	return newStringFormulaArg(result)
}

// ISNONTEXT function function tests if a supplied value is text. If not, the
// function returns TRUE; If the supplied value is text, the function returns
// FALSE. The syntax of the function is:
//
//    ISNONTEXT(value)
//
func (fn *formulaFuncs) ISNONTEXT(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "ISNONTEXT requires 1 argument")
	}
	token := argsList.Front().Value.(formulaArg)
	result := "TRUE"
	if token.Type == ArgString && token.String != "" {
		result = "FALSE"
	}
	return newStringFormulaArg(result)
}

// ISNUMBER function function tests if a supplied value is a number. If so,
// the function returns TRUE; Otherwise it returns FALSE. The syntax of the
// function is:
//
//    ISNUMBER(value)
//
func (fn *formulaFuncs) ISNUMBER(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "ISNUMBER requires 1 argument")
	}
	token, result := argsList.Front().Value.(formulaArg), false
	if token.Type == ArgString && token.String != "" {
		if _, err := strconv.Atoi(token.String); err == nil {
			result = true
		}
	}
	return newBoolFormulaArg(result)
}

// ISODD function tests if a supplied number (or numeric expression) evaluates
// to an odd number, and if so, returns TRUE; Otherwise, the function returns
// FALSE. The syntax of the function is:
//
//    ISODD(value)
//
func (fn *formulaFuncs) ISODD(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "ISODD requires 1 argument")
	}
	var (
		token   = argsList.Front().Value.(formulaArg)
		result  = "FALSE"
		numeric int
		err     error
	)
	if token.Type == ArgString {
		if numeric, err = strconv.Atoi(token.String); err != nil {
			return newErrorFormulaArg(formulaErrorVALUE, err.Error())
		}
		if numeric != numeric/2*2 {
			return newStringFormulaArg("TRUE")
		}
	}
	return newStringFormulaArg(result)
}

// ISTEXT function tests if a supplied value is text, and if so, returns TRUE;
// Otherwise, the function returns FALSE. The syntax of the function is:
//
//    ISTEXT(value)
//
func (fn *formulaFuncs) ISTEXT(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "ISTEXT requires 1 argument")
	}
	token := argsList.Front().Value.(formulaArg)
	if token.ToNumber().Type != ArgError {
		return newBoolFormulaArg(false)
	}
	return newBoolFormulaArg(token.Type == ArgString)
}

// N function converts data into a numeric value. The syntax of the function
// is:
//
//    N(value)
//
func (fn *formulaFuncs) N(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "N requires 1 argument")
	}
	token, num := argsList.Front().Value.(formulaArg), 0.0
	if token.Type == ArgError {
		return token
	}
	if arg := token.ToNumber(); arg.Type == ArgNumber {
		num = arg.Number
	}
	if token.Value() == "TRUE" {
		num = 1
	}
	return newNumberFormulaArg(num)
}

// NA function returns the Excel #N/A error. This error message has the
// meaning 'value not available' and is produced when an Excel Formula is
// unable to find a value that it needs. The syntax of the function is:
//
//    NA()
//
func (fn *formulaFuncs) NA(argsList *list.List) formulaArg {
	if argsList.Len() != 0 {
		return newErrorFormulaArg(formulaErrorVALUE, "NA accepts no arguments")
	}
	return newErrorFormulaArg(formulaErrorNA, formulaErrorNA)
}

// SHEET function returns the Sheet number for a specified reference. The
// syntax of the function is:
//
//    SHEET()
//
func (fn *formulaFuncs) SHEET(argsList *list.List) formulaArg {
	if argsList.Len() != 0 {
		return newErrorFormulaArg(formulaErrorVALUE, "SHEET accepts no arguments")
	}
	return newNumberFormulaArg(float64(fn.f.GetSheetIndex(fn.sheet) + 1))
}

// T function tests if a supplied value is text and if so, returns the
// supplied text; Otherwise, the function returns an empty text string. The
// syntax of the function is:
//
//    T(value)
//
func (fn *formulaFuncs) T(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "T requires 1 argument")
	}
	token := argsList.Front().Value.(formulaArg)
	if token.Type == ArgError {
		return token
	}
	if token.Type == ArgNumber {
		return newStringFormulaArg("")
	}
	return newStringFormulaArg(token.Value())
}

// Logical Functions

// AND function tests a number of supplied conditions and returns TRUE or
// FALSE. The syntax of the function is:
//
//    AND(logical_test1,[logical_test2],...)
//
func (fn *formulaFuncs) AND(argsList *list.List) formulaArg {
	if argsList.Len() == 0 {
		return newErrorFormulaArg(formulaErrorVALUE, "AND requires at least 1 argument")
	}
	if argsList.Len() > 30 {
		return newErrorFormulaArg(formulaErrorVALUE, "AND accepts at most 30 arguments")
	}
	var (
		and = true
		val float64
		err error
	)
	for arg := argsList.Front(); arg != nil; arg = arg.Next() {
		token := arg.Value.(formulaArg)
		switch token.Type {
		case ArgUnknown:
			continue
		case ArgString:
			if token.String == "TRUE" {
				continue
			}
			if token.String == "FALSE" {
				return newStringFormulaArg(token.String)
			}
			if val, err = strconv.ParseFloat(token.String, 64); err != nil {
				return newErrorFormulaArg(formulaErrorVALUE, err.Error())
			}
			and = and && (val != 0)
		case ArgMatrix:
			// TODO
			return newErrorFormulaArg(formulaErrorVALUE, formulaErrorVALUE)
		}
	}
	return newBoolFormulaArg(and)
}

// FALSE function function returns the logical value FALSE. The syntax of the
// function is:
//
//    FALSE()
//
func (fn *formulaFuncs) FALSE(argsList *list.List) formulaArg {
	if argsList.Len() != 0 {
		return newErrorFormulaArg(formulaErrorVALUE, "FALSE takes no arguments")
	}
	return newBoolFormulaArg(false)
}

// IFERROR function receives two values (or expressions) and tests if the
// first of these evaluates to an error. The syntax of the function is:
//
//    IFERROR(value,value_if_error)
//
func (fn *formulaFuncs) IFERROR(argsList *list.List) formulaArg {
	if argsList.Len() != 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "IFERROR requires 2 arguments")
	}
	value := argsList.Front().Value.(formulaArg)
	if value.Type != ArgError {
		if value.Type == ArgEmpty {
			return newNumberFormulaArg(0)
		}
		return value
	}
	return argsList.Back().Value.(formulaArg)
}

// NOT function returns the opposite to a supplied logical value. The syntax
// of the function is:
//
//    NOT(logical)
//
func (fn *formulaFuncs) NOT(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "NOT requires 1 argument")
	}
	token := argsList.Front().Value.(formulaArg)
	switch token.Type {
	case ArgString, ArgList:
		if strings.ToUpper(token.String) == "TRUE" {
			return newBoolFormulaArg(false)
		}
		if strings.ToUpper(token.String) == "FALSE" {
			return newBoolFormulaArg(true)
		}
	case ArgNumber:
		return newBoolFormulaArg(!(token.Number != 0))
	case ArgError:

		return token
	}
	return newErrorFormulaArg(formulaErrorVALUE, "NOT expects 1 boolean or numeric argument")
}

// OR function tests a number of supplied conditions and returns either TRUE
// or FALSE. The syntax of the function is:
//
//    OR(logical_test1,[logical_test2],...)
//
func (fn *formulaFuncs) OR(argsList *list.List) formulaArg {
	if argsList.Len() == 0 {
		return newErrorFormulaArg(formulaErrorVALUE, "OR requires at least 1 argument")
	}
	if argsList.Len() > 30 {
		return newErrorFormulaArg(formulaErrorVALUE, "OR accepts at most 30 arguments")
	}
	var (
		or  bool
		val float64
		err error
	)
	for arg := argsList.Front(); arg != nil; arg = arg.Next() {
		token := arg.Value.(formulaArg)
		switch token.Type {
		case ArgUnknown:
			continue
		case ArgString:
			if token.String == "FALSE" {
				continue
			}
			if token.String == "TRUE" {
				or = true
				continue
			}
			if val, err = strconv.ParseFloat(token.String, 64); err != nil {
				return newErrorFormulaArg(formulaErrorVALUE, err.Error())
			}
			or = val != 0
		case ArgMatrix:
			// TODO
			return newErrorFormulaArg(formulaErrorVALUE, formulaErrorVALUE)
		}
	}
	return newStringFormulaArg(strings.ToUpper(strconv.FormatBool(or)))
}

// TRUE function returns the logical value TRUE. The syntax of the function
// is:
//
//    TRUE()
//
func (fn *formulaFuncs) TRUE(argsList *list.List) formulaArg {
	if argsList.Len() != 0 {
		return newErrorFormulaArg(formulaErrorVALUE, "TRUE takes no arguments")
	}
	return newBoolFormulaArg(true)
}

// Date and Time Functions

// DATE returns a date, from a user-supplied year, month and day. The syntax
// of the function is:
//
//    DATE(year,month,day)
//
func (fn *formulaFuncs) DATE(argsList *list.List) formulaArg {
	if argsList.Len() != 3 {
		return newErrorFormulaArg(formulaErrorVALUE, "DATE requires 3 number arguments")
	}
	year := argsList.Front().Value.(formulaArg).ToNumber()
	month := argsList.Front().Next().Value.(formulaArg).ToNumber()
	day := argsList.Back().Value.(formulaArg).ToNumber()
	if year.Type != ArgNumber || month.Type != ArgNumber || day.Type != ArgNumber {
		return newErrorFormulaArg(formulaErrorVALUE, "DATE requires 3 number arguments")
	}
	d := makeDate(int(year.Number), time.Month(month.Number), int(day.Number))
	return newStringFormulaArg(timeFromExcelTime(daysBetween(excelMinTime1900.Unix(), d)+1, false).String())
}

// DATEDIF function calculates the number of days, months, or years between
// two dates. The syntax of the function is:
//
//    DATEDIF(start_date,end_date,unit)
//
func (fn *formulaFuncs) DATEDIF(argsList *list.List) formulaArg {
	if argsList.Len() != 3 {
		return newErrorFormulaArg(formulaErrorVALUE, "DATEDIF requires 3 number arguments")
	}
	startArg, endArg := argsList.Front().Value.(formulaArg).ToNumber(), argsList.Front().Next().Value.(formulaArg).ToNumber()
	if startArg.Type != ArgNumber || endArg.Type != ArgNumber {
		return startArg
	}
	if startArg.Number > endArg.Number {
		return newErrorFormulaArg(formulaErrorNUM, "start_date > end_date")
	}
	if startArg.Number == endArg.Number {
		return newNumberFormulaArg(0)
	}
	unit := strings.ToLower(argsList.Back().Value.(formulaArg).Value())
	startDate, endDate := timeFromExcelTime(startArg.Number, false), timeFromExcelTime(endArg.Number, false)
	sy, smm, sd := startDate.Date()
	ey, emm, ed := endDate.Date()
	sm, em, diff := int(smm), int(emm), 0.0
	switch unit {
	case "d":
		return newNumberFormulaArg(endArg.Number - startArg.Number)
	case "y":
		diff = float64(ey - sy)
		if em < sm || (em == sm && ed < sd) {
			diff--
		}
	case "m":
		ydiff := ey - sy
		mdiff := em - sm
		if ed < sd {
			mdiff--
		}
		if mdiff < 0 {
			ydiff--
			mdiff += 12
		}
		diff = float64(ydiff*12 + mdiff)
	case "md":
		smMD := em
		if ed < sd {
			smMD--
		}
		diff = endArg.Number - daysBetween(excelMinTime1900.Unix(), makeDate(ey, time.Month(smMD), sd)) - 1
	case "ym":
		diff = float64(em - sm)
		if ed < sd {
			diff--
		}
		if diff < 0 {
			diff += 12
		}
	case "yd":
		syYD := sy
		if em < sm || (em == sm && ed < sd) {
			syYD++
		}
		s := daysBetween(excelMinTime1900.Unix(), makeDate(syYD, time.Month(em), ed))
		e := daysBetween(excelMinTime1900.Unix(), makeDate(sy, time.Month(sm), sd))
		diff = s - e
	default:
		return newErrorFormulaArg(formulaErrorVALUE, "DATEDIF has invalid unit")
	}
	return newNumberFormulaArg(diff)
}

// NOW function returns the current date and time. The function receives no
// arguments and therefore. The syntax of the function is:
//
//    NOW()
//
func (fn *formulaFuncs) NOW(argsList *list.List) formulaArg {
	if argsList.Len() != 0 {
		return newErrorFormulaArg(formulaErrorVALUE, "NOW accepts no arguments")
	}
	now := time.Now()
	_, offset := now.Zone()
	return newNumberFormulaArg(25569.0 + float64(now.Unix()+int64(offset))/86400)
}

// TODAY function returns the current date. The function has no arguments and
// therefore. The syntax of the function is:
//
//    TODAY()
//
func (fn *formulaFuncs) TODAY(argsList *list.List) formulaArg {
	if argsList.Len() != 0 {
		return newErrorFormulaArg(formulaErrorVALUE, "TODAY accepts no arguments")
	}
	now := time.Now()
	_, offset := now.Zone()
	return newNumberFormulaArg(daysBetween(excelMinTime1900.Unix(), now.Unix()+int64(offset)) + 1)
}

// makeDate return date as a Unix time, the number of seconds elapsed since
// January 1, 1970 UTC.
func makeDate(y int, m time.Month, d int) int64 {
	if y == 1900 && int(m) <= 2 {
		d--
	}
	date := time.Date(y, m, d, 0, 0, 0, 0, time.UTC)
	return date.Unix()
}

// daysBetween return time interval of the given start timestamp and end
// timestamp.
func daysBetween(startDate, endDate int64) float64 {
	return float64(int(0.5 + float64((endDate-startDate)/86400)))
}

// Text Functions

// CHAR function returns the character relating to a supplied character set
// number (from 1 to 255). syntax of the function is:
//
//    CHAR(number)
//
func (fn *formulaFuncs) CHAR(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "CHAR requires 1 argument")
	}
	arg := argsList.Front().Value.(formulaArg).ToNumber()
	if arg.Type != ArgNumber {
		return arg
	}
	num := int(arg.Number)
	if num < 0 || num > 255 {
		return newErrorFormulaArg(formulaErrorVALUE, formulaErrorVALUE)
	}
	return newStringFormulaArg(fmt.Sprintf("%c", num))
}

// CLEAN removes all non-printable characters from a supplied text string. The
// syntax of the function is:
//
//    CLEAN(text)
//
func (fn *formulaFuncs) CLEAN(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "CLEAN requires 1 argument")
	}
	b := bytes.Buffer{}
	for _, c := range argsList.Front().Value.(formulaArg).String {
		if c > 31 {
			b.WriteRune(c)
		}
	}
	return newStringFormulaArg(b.String())
}

// CODE function converts the first character of a supplied text string into
// the associated numeric character set code used by your computer. The
// syntax of the function is:
//
//    CODE(text)
//
func (fn *formulaFuncs) CODE(argsList *list.List) formulaArg {
	return fn.code("CODE", argsList)
}

// code is an implementation of the formula function CODE and UNICODE.
func (fn *formulaFuncs) code(name string, argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, fmt.Sprintf("%s requires 1 argument", name))
	}
	text := argsList.Front().Value.(formulaArg).Value()
	if len(text) == 0 {
		if name == "CODE" {
			return newNumberFormulaArg(0)
		}
		return newErrorFormulaArg(formulaErrorVALUE, formulaErrorVALUE)
	}
	return newNumberFormulaArg(float64(text[0]))
}

// CONCAT function joins together a series of supplied text strings into one
// combined text string.
//
//    CONCAT(text1,[text2],...)
//
func (fn *formulaFuncs) CONCAT(argsList *list.List) formulaArg {
	return fn.concat("CONCAT", argsList)
}

// CONCATENATE function joins together a series of supplied text strings into
// one combined text string.
//
//    CONCATENATE(text1,[text2],...)
//
func (fn *formulaFuncs) CONCATENATE(argsList *list.List) formulaArg {
	return fn.concat("CONCATENATE", argsList)
}

// concat is an implementation of the formula function CONCAT and CONCATENATE.
func (fn *formulaFuncs) concat(name string, argsList *list.List) formulaArg {
	buf := bytes.Buffer{}
	for arg := argsList.Front(); arg != nil; arg = arg.Next() {
		token := arg.Value.(formulaArg)
		switch token.Type {
		case ArgString:
			buf.WriteString(token.String)
		case ArgNumber:
			if token.Boolean {
				if token.Number == 0 {
					buf.WriteString("FALSE")
				} else {
					buf.WriteString("TRUE")
				}
			} else {
				buf.WriteString(token.Value())
			}
		default:
			return newErrorFormulaArg(formulaErrorVALUE, fmt.Sprintf("%s requires arguments to be strings", name))
		}
	}
	return newStringFormulaArg(buf.String())
}

// EXACT function tests if two supplied text strings or values are exactly
// equal and if so, returns TRUE; Otherwise, the function returns FALSE. The
// function is case-sensitive. The syntax of the function is:
//
//    EXACT(text1,text2)
//
func (fn *formulaFuncs) EXACT(argsList *list.List) formulaArg {
	if argsList.Len() != 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "EXACT requires 2 arguments")
	}
	text1 := argsList.Front().Value.(formulaArg).Value()
	text2 := argsList.Back().Value.(formulaArg).Value()
	return newBoolFormulaArg(text1 == text2)
}

// FIXED function rounds a supplied number to a specified number of decimal
// places and then converts this into text. The syntax of the function is:
//
//    FIXED(number,[decimals],[no_commas])
//
func (fn *formulaFuncs) FIXED(argsList *list.List) formulaArg {
	if argsList.Len() < 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "FIXED requires at least 1 argument")
	}
	if argsList.Len() > 3 {
		return newErrorFormulaArg(formulaErrorVALUE, "FIXED allows at most 3 arguments")
	}
	numArg := argsList.Front().Value.(formulaArg).ToNumber()
	if numArg.Type != ArgNumber {
		return numArg
	}
	precision, decimals, noCommas := 0, 0, false
	s := strings.Split(argsList.Front().Value.(formulaArg).Value(), ".")
	if argsList.Len() == 1 && len(s) == 2 {
		precision = len(s[1])
		decimals = len(s[1])
	}
	if argsList.Len() >= 2 {
		decimalsArg := argsList.Front().Next().Value.(formulaArg).ToNumber()
		if decimalsArg.Type != ArgNumber {
			return decimalsArg
		}
		decimals = int(decimalsArg.Number)
	}
	if argsList.Len() == 3 {
		noCommasArg := argsList.Back().Value.(formulaArg).ToBool()
		if noCommasArg.Type == ArgError {
			return noCommasArg
		}
		noCommas = noCommasArg.Boolean
	}
	n := math.Pow(10, float64(decimals))
	r := numArg.Number * n
	fixed := float64(int(r+math.Copysign(0.5, r))) / n
	if decimals > 0 {
		precision = decimals
	}
	if noCommas {
		return newStringFormulaArg(fmt.Sprintf(fmt.Sprintf("%%.%df", precision), fixed))
	}
	p := message.NewPrinter(language.English)
	return newStringFormulaArg(p.Sprintf(fmt.Sprintf("%%.%df", precision), fixed))
}

// FIND function returns the position of a specified character or sub-string
// within a supplied text string. The function is case-sensitive. The syntax
// of the function is:
//
//    FIND(find_text,within_text,[start_num])
//
func (fn *formulaFuncs) FIND(argsList *list.List) formulaArg {
	return fn.find("FIND", argsList)
}

// FINDB counts each double-byte character as 2 when you have enabled the
// editing of a language that supports DBCS and then set it as the default
// language. Otherwise, FINDB counts each character as 1. The syntax of the
// function is:
//
//    FINDB(find_text,within_text,[start_num])
//
func (fn *formulaFuncs) FINDB(argsList *list.List) formulaArg {
	return fn.find("FINDB", argsList)
}

// find is an implementation of the formula function FIND and FINDB.
func (fn *formulaFuncs) find(name string, argsList *list.List) formulaArg {
	if argsList.Len() < 2 {
		return newErrorFormulaArg(formulaErrorVALUE, fmt.Sprintf("%s requires at least 2 arguments", name))
	}
	if argsList.Len() > 3 {
		return newErrorFormulaArg(formulaErrorVALUE, fmt.Sprintf("%s allows at most 3 arguments", name))
	}
	findText := argsList.Front().Value.(formulaArg).Value()
	withinText := argsList.Front().Next().Value.(formulaArg).Value()
	startNum, result := 1, 1
	if argsList.Len() == 3 {
		numArg := argsList.Back().Value.(formulaArg).ToNumber()
		if numArg.Type != ArgNumber {
			return numArg
		}
		if numArg.Number < 0 {
			return newErrorFormulaArg(formulaErrorVALUE, formulaErrorVALUE)
		}
		startNum = int(numArg.Number)
	}
	if findText == "" {
		return newNumberFormulaArg(float64(startNum))
	}
	for idx := range withinText {
		if result < startNum {
			result++
		}
		if strings.Index(withinText[idx:], findText) == 0 {
			return newNumberFormulaArg(float64(result))
		}
		result++
	}
	return newErrorFormulaArg(formulaErrorVALUE, formulaErrorVALUE)
}

// LEFT function returns a specified number of characters from the start of a
// supplied text string. The syntax of the function is:
//
//    LEFT(text,[num_chars])
//
func (fn *formulaFuncs) LEFT(argsList *list.List) formulaArg {
	return fn.leftRight("LEFT", argsList)
}

// LEFTB returns the first character or characters in a text string, based on
// the number of bytes you specify. The syntax of the function is:
//
//    LEFTB(text,[num_bytes])
//
func (fn *formulaFuncs) LEFTB(argsList *list.List) formulaArg {
	return fn.leftRight("LEFTB", argsList)
}

// leftRight is an implementation of the formula function LEFT, LEFTB, RIGHT,
// RIGHTB. TODO: support DBCS include Japanese, Chinese (Simplified), Chinese
// (Traditional), and Korean.
func (fn *formulaFuncs) leftRight(name string, argsList *list.List) formulaArg {
	if argsList.Len() < 1 {
		return newErrorFormulaArg(formulaErrorVALUE, fmt.Sprintf("%s requires at least 1 argument", name))
	}
	if argsList.Len() > 2 {
		return newErrorFormulaArg(formulaErrorVALUE, fmt.Sprintf("%s allows at most 2 arguments", name))
	}
	text, numChars := argsList.Front().Value.(formulaArg).Value(), 1
	if argsList.Len() == 2 {
		numArg := argsList.Back().Value.(formulaArg).ToNumber()
		if numArg.Type != ArgNumber {
			return numArg
		}
		if numArg.Number < 0 {
			return newErrorFormulaArg(formulaErrorVALUE, formulaErrorVALUE)
		}
		numChars = int(numArg.Number)
	}
	if len(text) > numChars {
		if name == "LEFT" || name == "LEFTB" {
			return newStringFormulaArg(text[:numChars])
		}
		return newStringFormulaArg(text[len(text)-numChars:])
	}
	return newStringFormulaArg(text)
}

// LEN returns the length of a supplied text string. The syntax of the
// function is:
//
//    LEN(text)
//
func (fn *formulaFuncs) LEN(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "LEN requires 1 string argument")
	}
	return newStringFormulaArg(strconv.Itoa(len(argsList.Front().Value.(formulaArg).String)))
}

// LENB returns the number of bytes used to represent the characters in a text
// string. LENB counts 2 bytes per character only when a DBCS language is set
// as the default language. Otherwise LENB behaves the same as LEN, counting
// 1 byte per character. The syntax of the function is:
//
//    LENB(text)
//
// TODO: the languages that support DBCS include Japanese, Chinese
// (Simplified), Chinese (Traditional), and Korean.
func (fn *formulaFuncs) LENB(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "LENB requires 1 string argument")
	}
	return newStringFormulaArg(strconv.Itoa(len(argsList.Front().Value.(formulaArg).String)))
}

// LOWER converts all characters in a supplied text string to lower case. The
// syntax of the function is:
//
//    LOWER(text)
//
func (fn *formulaFuncs) LOWER(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "LOWER requires 1 argument")
	}
	return newStringFormulaArg(strings.ToLower(argsList.Front().Value.(formulaArg).String))
}

// MID function returns a specified number of characters from the middle of a
// supplied text string. The syntax of the function is:
//
//    MID(text,start_num,num_chars)
//
func (fn *formulaFuncs) MID(argsList *list.List) formulaArg {
	return fn.mid("MID", argsList)
}

// MIDB returns a specific number of characters from a text string, starting
// at the position you specify, based on the number of bytes you specify. The
// syntax of the function is:
//
//    MID(text,start_num,num_chars)
//
func (fn *formulaFuncs) MIDB(argsList *list.List) formulaArg {
	return fn.mid("MIDB", argsList)
}

// mid is an implementation of the formula function MID and MIDB. TODO:
// support DBCS include Japanese, Chinese (Simplified), Chinese
// (Traditional), and Korean.
func (fn *formulaFuncs) mid(name string, argsList *list.List) formulaArg {
	if argsList.Len() != 3 {
		return newErrorFormulaArg(formulaErrorVALUE, fmt.Sprintf("%s requires 3 arguments", name))
	}
	text := argsList.Front().Value.(formulaArg).Value()
	startNumArg, numCharsArg := argsList.Front().Next().Value.(formulaArg).ToNumber(), argsList.Front().Next().Next().Value.(formulaArg).ToNumber()
	if startNumArg.Type != ArgNumber {
		return startNumArg
	}
	if numCharsArg.Type != ArgNumber {
		return numCharsArg
	}
	startNum := int(startNumArg.Number)
	if startNum < 0 {
		return newErrorFormulaArg(formulaErrorVALUE, formulaErrorVALUE)
	}
	textLen := len(text)
	if startNum > textLen {
		return newStringFormulaArg("")
	}
	startNum--
	endNum := startNum + int(numCharsArg.Number)
	if endNum > textLen+1 {
		return newStringFormulaArg(text[startNum:])
	}
	return newStringFormulaArg(text[startNum:endNum])
}

// PROPER converts all characters in a supplied text string to proper case
// (i.e. all letters that do not immediately follow another letter are set to
// upper case and all other characters are lower case). The syntax of the
// function is:
//
//    PROPER(text)
//
func (fn *formulaFuncs) PROPER(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "PROPER requires 1 argument")
	}
	buf := bytes.Buffer{}
	isLetter := false
	for _, char := range argsList.Front().Value.(formulaArg).String {
		if !isLetter && unicode.IsLetter(char) {
			buf.WriteRune(unicode.ToUpper(char))
		} else {
			buf.WriteRune(unicode.ToLower(char))
		}
		isLetter = unicode.IsLetter(char)
	}
	return newStringFormulaArg(buf.String())
}

// REPLACE function replaces all or part of a text string with another string.
// The syntax of the function is:
//
//    REPLACE(old_text,start_num,num_chars,new_text)
//
func (fn *formulaFuncs) REPLACE(argsList *list.List) formulaArg {
	return fn.replace("REPLACE", argsList)
}

// REPLACEB replaces part of a text string, based on the number of bytes you
// specify, with a different text string.
//
//    REPLACEB(old_text,start_num,num_chars,new_text)
//
func (fn *formulaFuncs) REPLACEB(argsList *list.List) formulaArg {
	return fn.replace("REPLACEB", argsList)
}

// replace is an implementation of the formula function REPLACE and REPLACEB.
// TODO: support DBCS include Japanese, Chinese (Simplified), Chinese
// (Traditional), and Korean.
func (fn *formulaFuncs) replace(name string, argsList *list.List) formulaArg {
	if argsList.Len() != 4 {
		return newErrorFormulaArg(formulaErrorVALUE, fmt.Sprintf("%s requires 4 arguments", name))
	}
	oldText, newText := argsList.Front().Value.(formulaArg).Value(), argsList.Back().Value.(formulaArg).Value()
	startNumArg, numCharsArg := argsList.Front().Next().Value.(formulaArg).ToNumber(), argsList.Front().Next().Next().Value.(formulaArg).ToNumber()
	if startNumArg.Type != ArgNumber {
		return startNumArg
	}
	if numCharsArg.Type != ArgNumber {
		return numCharsArg
	}
	oldTextLen, startIdx := len(oldText), int(startNumArg.Number)
	if startIdx > oldTextLen {
		startIdx = oldTextLen + 1
	}
	endIdx := startIdx + int(numCharsArg.Number)
	if endIdx > oldTextLen {
		endIdx = oldTextLen + 1
	}
	if startIdx < 1 || endIdx < 1 {
		return newErrorFormulaArg(formulaErrorVALUE, formulaErrorVALUE)
	}
	result := oldText[:startIdx-1] + newText + oldText[endIdx-1:]
	return newStringFormulaArg(result)
}

// REPT function returns a supplied text string, repeated a specified number
// of times. The syntax of the function is:
//
//    REPT(text,number_times)
//
func (fn *formulaFuncs) REPT(argsList *list.List) formulaArg {
	if argsList.Len() != 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "REPT requires 2 arguments")
	}
	text := argsList.Front().Value.(formulaArg)
	if text.Type != ArgString {
		return newErrorFormulaArg(formulaErrorVALUE, "REPT requires first argument to be a string")
	}
	times := argsList.Back().Value.(formulaArg).ToNumber()
	if times.Type != ArgNumber {
		return newErrorFormulaArg(formulaErrorVALUE, "REPT requires second argument to be a number")
	}
	if times.Number < 0 {
		return newErrorFormulaArg(formulaErrorVALUE, "REPT requires second argument to be >= 0")
	}
	if times.Number == 0 {
		return newStringFormulaArg("")
	}
	buf := bytes.Buffer{}
	for i := 0; i < int(times.Number); i++ {
		buf.WriteString(text.String)
	}
	return newStringFormulaArg(buf.String())
}

// RIGHT function returns a specified number of characters from the end of a
// supplied text string. The syntax of the function is:
//
//    RIGHT(text,[num_chars])
//
func (fn *formulaFuncs) RIGHT(argsList *list.List) formulaArg {
	return fn.leftRight("RIGHT", argsList)
}

// RIGHTB returns the last character or characters in a text string, based on
// the number of bytes you specify. The syntax of the function is:
//
//    RIGHTB(text,[num_bytes])
//
func (fn *formulaFuncs) RIGHTB(argsList *list.List) formulaArg {
	return fn.leftRight("RIGHTB", argsList)
}

// SUBSTITUTE function replaces one or more instances of a given text string,
// within an original text string. The syntax of the function is:
//
//    SUBSTITUTE(text,old_text,new_text,[instance_num])
//
func (fn *formulaFuncs) SUBSTITUTE(argsList *list.List) formulaArg {
	if argsList.Len() != 3 && argsList.Len() != 4 {
		return newErrorFormulaArg(formulaErrorVALUE, "SUBSTITUTE requires 3 or 4 arguments")
	}
	text, oldText := argsList.Front().Value.(formulaArg), argsList.Front().Next().Value.(formulaArg)
	newText, instanceNum := argsList.Front().Next().Next().Value.(formulaArg), 0
	if argsList.Len() == 3 {
		return newStringFormulaArg(strings.Replace(text.Value(), oldText.Value(), newText.Value(), -1))
	}
	instanceNumArg := argsList.Back().Value.(formulaArg).ToNumber()
	if instanceNumArg.Type != ArgNumber {
		return instanceNumArg
	}
	instanceNum = int(instanceNumArg.Number)
	if instanceNum < 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "instance_num should be > 0")
	}
	str, oldTextLen, count, chars, pos := text.Value(), len(oldText.Value()), instanceNum, 0, -1
	for {
		count--
		index := strings.Index(str, oldText.Value())
		if index == -1 {
			pos = -1
			break
		} else {
			pos = index + chars
			if count == 0 {
				break
			}
			idx := oldTextLen + index
			chars += idx
			str = str[idx:]
		}
	}
	if pos == -1 {
		return newStringFormulaArg(text.Value())
	}
	pre, post := text.Value()[:pos], text.Value()[pos+oldTextLen:]
	return newStringFormulaArg(pre + newText.Value() + post)
}

// TRIM removes extra spaces (i.e. all spaces except for single spaces between
// words or characters) from a supplied text string. The syntax of the
// function is:
//
//    TRIM(text)
//
func (fn *formulaFuncs) TRIM(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "TRIM requires 1 argument")
	}
	return newStringFormulaArg(strings.TrimSpace(argsList.Front().Value.(formulaArg).String))
}

// UNICHAR returns the Unicode character that is referenced by the given
// numeric value. The syntax of the function is:
//
//    UNICHAR(number)
//
func (fn *formulaFuncs) UNICHAR(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "UNICHAR requires 1 argument")
	}
	numArg := argsList.Front().Value.(formulaArg).ToNumber()
	if numArg.Type != ArgNumber {
		return numArg
	}
	if numArg.Number <= 0 || numArg.Number > 55295 {
		return newErrorFormulaArg(formulaErrorVALUE, formulaErrorVALUE)
	}
	return newStringFormulaArg(string(rune(numArg.Number)))
}

// UNICODE function returns the code point for the first character of a
// supplied text string. The syntax of the function is:
//
//    UNICODE(text)
//
func (fn *formulaFuncs) UNICODE(argsList *list.List) formulaArg {
	return fn.code("UNICODE", argsList)
}

// UPPER converts all characters in a supplied text string to upper case. The
// syntax of the function is:
//
//    UPPER(text)
//
func (fn *formulaFuncs) UPPER(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "UPPER requires 1 argument")
	}
	return newStringFormulaArg(strings.ToUpper(argsList.Front().Value.(formulaArg).String))
}

// Conditional Functions

// IF function tests a supplied condition and returns one result if the
// condition evaluates to TRUE, and another result if the condition evaluates
// to FALSE. The syntax of the function is:
//
//    IF(logical_test,value_if_true,value_if_false)
//
func (fn *formulaFuncs) IF(argsList *list.List) formulaArg {
	if argsList.Len() == 0 {
		return newErrorFormulaArg(formulaErrorVALUE, "IF requires at least 1 argument")
	}
	if argsList.Len() > 3 {
		return newErrorFormulaArg(formulaErrorVALUE, "IF accepts at most 3 arguments")
	}
	token := argsList.Front().Value.(formulaArg)
	var (
		cond   bool
		err    error
		result string
	)
	switch token.Type {
	case ArgString:
		if cond, err = strconv.ParseBool(token.String); err != nil {
			return newErrorFormulaArg(formulaErrorVALUE, err.Error())
		}
		if argsList.Len() == 1 {
			return newBoolFormulaArg(cond)
		}
		if cond {
			return newStringFormulaArg(argsList.Front().Next().Value.(formulaArg).String)
		}
		if argsList.Len() == 3 {
			result = argsList.Back().Value.(formulaArg).String
		}
	}
	return newStringFormulaArg(result)
}

// Lookup and Reference Functions

// CHOOSE function returns a value from an array, that corresponds to a
// supplied index number (position). The syntax of the function is:
//
//    CHOOSE(index_num,value1,[value2],...)
//
func (fn *formulaFuncs) CHOOSE(argsList *list.List) formulaArg {
	if argsList.Len() < 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "CHOOSE requires 2 arguments")
	}
	idx, err := strconv.Atoi(argsList.Front().Value.(formulaArg).String)
	if err != nil {
		return newErrorFormulaArg(formulaErrorVALUE, "CHOOSE requires first argument of type number")
	}
	if argsList.Len() <= idx {
		return newErrorFormulaArg(formulaErrorVALUE, "index_num should be <= to the number of values")
	}
	arg := argsList.Front()
	for i := 0; i < idx; i++ {
		arg = arg.Next()
	}
	var result formulaArg
	switch arg.Value.(formulaArg).Type {
	case ArgString:
		result = newStringFormulaArg(arg.Value.(formulaArg).String)
	case ArgMatrix:
		result = newMatrixFormulaArg(arg.Value.(formulaArg).Matrix)
	}
	return result
}

// deepMatchRune finds whether the text deep matches/satisfies the pattern
// string.
func deepMatchRune(str, pattern []rune, simple bool) bool {
	for len(pattern) > 0 {
		switch pattern[0] {
		default:
			if len(str) == 0 || str[0] != pattern[0] {
				return false
			}
		case '?':
			if len(str) == 0 && !simple {
				return false
			}
		case '*':
			return deepMatchRune(str, pattern[1:], simple) ||
				(len(str) > 0 && deepMatchRune(str[1:], pattern, simple))
		}
		str = str[1:]
		pattern = pattern[1:]
	}
	return len(str) == 0 && len(pattern) == 0
}

// matchPattern finds whether the text matches or satisfies the pattern
// string. The pattern supports '*' and '?' wildcards in the pattern string.
func matchPattern(pattern, name string) (matched bool) {
	if pattern == "" {
		return name == pattern
	}
	if pattern == "*" {
		return true
	}
	rname, rpattern := make([]rune, 0, len(name)), make([]rune, 0, len(pattern))
	for _, r := range name {
		rname = append(rname, r)
	}
	for _, r := range pattern {
		rpattern = append(rpattern, r)
	}
	simple := false // Does extended wildcard '*' and '?' match.
	return deepMatchRune(rname, rpattern, simple)
}

// compareFormulaArg compares the left-hand sides and the right-hand sides
// formula arguments by given conditions such as case sensitive, if exact
// match, and make compare result as formula criteria condition type.
func compareFormulaArg(lhs, rhs formulaArg, caseSensitive, exactMatch bool) byte {
	if lhs.Type != rhs.Type {
		return criteriaErr
	}
	switch lhs.Type {
	case ArgNumber:
		if lhs.Number == rhs.Number {
			return criteriaEq
		}
		if lhs.Number < rhs.Number {
			return criteriaL
		}
		return criteriaG
	case ArgString:
		ls, rs := lhs.String, rhs.String
		if !caseSensitive {
			ls, rs = strings.ToLower(ls), strings.ToLower(rs)
		}
		if exactMatch {
			match := matchPattern(rs, ls)
			if match {
				return criteriaEq
			}
			return criteriaG
		}
		switch strings.Compare(ls, rs) {
		case 1:
			return criteriaG
		case -1:
			return criteriaL
		case 0:
			return criteriaEq
		}
		return criteriaErr
	case ArgEmpty:
		return criteriaEq
	case ArgList:
		return compareFormulaArgList(lhs, rhs, caseSensitive, exactMatch)
	case ArgMatrix:
		return compareFormulaArgMatrix(lhs, rhs, caseSensitive, exactMatch)
	}
	return criteriaErr
}

// compareFormulaArgList compares the left-hand sides and the right-hand sides
// list type formula arguments.
func compareFormulaArgList(lhs, rhs formulaArg, caseSensitive, exactMatch bool) byte {
	if len(lhs.List) < len(rhs.List) {
		return criteriaL
	}
	if len(lhs.List) > len(rhs.List) {
		return criteriaG
	}
	for arg := range lhs.List {
		criteria := compareFormulaArg(lhs.List[arg], rhs.List[arg], caseSensitive, exactMatch)
		if criteria != criteriaEq {
			return criteria
		}
	}
	return criteriaEq
}

// compareFormulaArgMatrix compares the left-hand sides and the right-hand sides
// matrix type formula arguments.
func compareFormulaArgMatrix(lhs, rhs formulaArg, caseSensitive, exactMatch bool) byte {
	if len(lhs.Matrix) < len(rhs.Matrix) {
		return criteriaL
	}
	if len(lhs.Matrix) > len(rhs.Matrix) {
		return criteriaG
	}
	for i := range lhs.Matrix {
		left := lhs.Matrix[i]
		right := lhs.Matrix[i]
		if len(left) < len(right) {
			return criteriaL
		}
		if len(left) > len(right) {
			return criteriaG
		}
		for arg := range left {
			criteria := compareFormulaArg(left[arg], right[arg], caseSensitive, exactMatch)
			if criteria != criteriaEq {
				return criteria
			}
		}
	}
	return criteriaEq
}

// COLUMN function returns the first column number within a supplied reference
// or the number of the current column. The syntax of the function is:
//
//    COLUMN([reference])
//
func (fn *formulaFuncs) COLUMN(argsList *list.List) formulaArg {
	if argsList.Len() > 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "COLUMN requires at most 1 argument")
	}
	if argsList.Len() == 1 {
		if argsList.Front().Value.(formulaArg).cellRanges != nil && argsList.Front().Value.(formulaArg).cellRanges.Len() > 0 {
			return newNumberFormulaArg(float64(argsList.Front().Value.(formulaArg).cellRanges.Front().Value.(cellRange).From.Col))
		}
		if argsList.Front().Value.(formulaArg).cellRefs != nil && argsList.Front().Value.(formulaArg).cellRefs.Len() > 0 {
			return newNumberFormulaArg(float64(argsList.Front().Value.(formulaArg).cellRefs.Front().Value.(cellRef).Col))
		}
		return newErrorFormulaArg(formulaErrorVALUE, "invalid reference")
	}
	col, _, _ := CellNameToCoordinates(fn.cell)
	return newNumberFormulaArg(float64(col))
}

// COLUMNS function receives an Excel range and returns the number of columns
// that are contained within the range. The syntax of the function is:
//
//    COLUMNS(array)
//
func (fn *formulaFuncs) COLUMNS(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "COLUMNS requires 1 argument")
	}
	var min, max int
	if argsList.Front().Value.(formulaArg).cellRanges != nil && argsList.Front().Value.(formulaArg).cellRanges.Len() > 0 {
		crs := argsList.Front().Value.(formulaArg).cellRanges
		for cr := crs.Front(); cr != nil; cr = cr.Next() {
			if min == 0 {
				min = cr.Value.(cellRange).From.Col
			}
			if min > cr.Value.(cellRange).From.Col {
				min = cr.Value.(cellRange).From.Col
			}
			if min > cr.Value.(cellRange).To.Col {
				min = cr.Value.(cellRange).To.Col
			}
			if max < cr.Value.(cellRange).To.Col {
				max = cr.Value.(cellRange).To.Col
			}
			if max < cr.Value.(cellRange).From.Col {
				max = cr.Value.(cellRange).From.Col
			}
		}
	}
	if argsList.Front().Value.(formulaArg).cellRefs != nil && argsList.Front().Value.(formulaArg).cellRefs.Len() > 0 {
		cr := argsList.Front().Value.(formulaArg).cellRefs
		for refs := cr.Front(); refs != nil; refs = refs.Next() {
			if min == 0 {
				min = refs.Value.(cellRef).Col
			}
			if min > refs.Value.(cellRef).Col {
				min = refs.Value.(cellRef).Col
			}
			if max < refs.Value.(cellRef).Col {
				max = refs.Value.(cellRef).Col
			}
		}
	}
	if max == TotalColumns {
		return newNumberFormulaArg(float64(TotalColumns))
	}
	result := max - min + 1
	if max == min {
		if min == 0 {
			return newErrorFormulaArg(formulaErrorVALUE, "invalid reference")
		}
		return newNumberFormulaArg(float64(1))
	}
	return newNumberFormulaArg(float64(result))
}

// HLOOKUP function 'looks up' a given value in the top row of a data array
// (or table), and returns the corresponding value from another row of the
// array. The syntax of the function is:
//
//    HLOOKUP(lookup_value,table_array,row_index_num,[range_lookup])
//
func (fn *formulaFuncs) HLOOKUP(argsList *list.List) formulaArg {
	if argsList.Len() < 3 {
		return newErrorFormulaArg(formulaErrorVALUE, "HLOOKUP requires at least 3 arguments")
	}
	if argsList.Len() > 4 {
		return newErrorFormulaArg(formulaErrorVALUE, "HLOOKUP requires at most 4 arguments")
	}
	lookupValue := argsList.Front().Value.(formulaArg)
	tableArray := argsList.Front().Next().Value.(formulaArg)
	if tableArray.Type != ArgMatrix {
		return newErrorFormulaArg(formulaErrorVALUE, "HLOOKUP requires second argument of table array")
	}
	rowArg := argsList.Front().Next().Next().Value.(formulaArg).ToNumber()
	if rowArg.Type != ArgNumber {
		return newErrorFormulaArg(formulaErrorVALUE, "HLOOKUP requires numeric row argument")
	}
	rowIdx, matchIdx, wasExact, exactMatch := int(rowArg.Number)-1, -1, false, false
	if argsList.Len() == 4 {
		rangeLookup := argsList.Back().Value.(formulaArg).ToBool()
		if rangeLookup.Type == ArgError {
			return newErrorFormulaArg(formulaErrorVALUE, rangeLookup.Error)
		}
		if rangeLookup.Number == 0 {
			exactMatch = true
		}
	}
	row := tableArray.Matrix[0]
	if exactMatch || len(tableArray.Matrix) == TotalRows {
	start:
		for idx, mtx := range row {
			lhs := mtx
			switch lookupValue.Type {
			case ArgNumber:
				if !lookupValue.Boolean {
					lhs = mtx.ToNumber()
					if lhs.Type == ArgError {
						lhs = mtx
					}
				}
			case ArgMatrix:
				lhs = tableArray
			}
			if compareFormulaArg(lhs, lookupValue, false, exactMatch) == criteriaEq {
				matchIdx = idx
				wasExact = true
				break start
			}
		}
	} else {
		matchIdx, wasExact = hlookupBinarySearch(row, lookupValue)
	}
	if matchIdx == -1 {
		return newErrorFormulaArg(formulaErrorNA, "HLOOKUP no result found")
	}
	if rowIdx < 0 || rowIdx >= len(tableArray.Matrix) {
		return newErrorFormulaArg(formulaErrorNA, "HLOOKUP has invalid row index")
	}
	row = tableArray.Matrix[rowIdx]
	if wasExact || !exactMatch {
		return row[matchIdx]
	}
	return newErrorFormulaArg(formulaErrorNA, "HLOOKUP no result found")
}

// VLOOKUP function 'looks up' a given value in the left-hand column of a
// data array (or table), and returns the corresponding value from another
// column of the array. The syntax of the function is:
//
//    VLOOKUP(lookup_value,table_array,col_index_num,[range_lookup])
//
func (fn *formulaFuncs) VLOOKUP(argsList *list.List) formulaArg {
	if argsList.Len() < 3 {
		return newErrorFormulaArg(formulaErrorVALUE, "VLOOKUP requires at least 3 arguments")
	}
	if argsList.Len() > 4 {
		return newErrorFormulaArg(formulaErrorVALUE, "VLOOKUP requires at most 4 arguments")
	}
	lookupValue := argsList.Front().Value.(formulaArg)
	tableArray := argsList.Front().Next().Value.(formulaArg)
	if tableArray.Type != ArgMatrix {
		return newErrorFormulaArg(formulaErrorVALUE, "VLOOKUP requires second argument of table array")
	}
	colIdx := argsList.Front().Next().Next().Value.(formulaArg).ToNumber()
	if colIdx.Type != ArgNumber {
		return newErrorFormulaArg(formulaErrorVALUE, "VLOOKUP requires numeric col argument")
	}
	col, matchIdx, wasExact, exactMatch := int(colIdx.Number)-1, -1, false, false
	if argsList.Len() == 4 {
		rangeLookup := argsList.Back().Value.(formulaArg).ToBool()
		if rangeLookup.Type == ArgError {
			return newErrorFormulaArg(formulaErrorVALUE, rangeLookup.Error)
		}
		if rangeLookup.Number == 0 {
			exactMatch = true
		}
	}
	if exactMatch || len(tableArray.Matrix) == TotalRows {
	start:
		for idx, mtx := range tableArray.Matrix {
			lhs := mtx[0]
			switch lookupValue.Type {
			case ArgNumber:
				if !lookupValue.Boolean {
					lhs = mtx[0].ToNumber()
					if lhs.Type == ArgError {
						lhs = mtx[0]
					}
				}
			case ArgMatrix:
				lhs = tableArray
			}
			if compareFormulaArg(lhs, lookupValue, false, exactMatch) == criteriaEq {
				matchIdx = idx
				wasExact = true
				break start
			}
		}
	} else {
		matchIdx, wasExact = vlookupBinarySearch(tableArray, lookupValue)
	}
	if matchIdx == -1 {
		return newErrorFormulaArg(formulaErrorNA, "VLOOKUP no result found")
	}
	mtx := tableArray.Matrix[matchIdx]
	if col < 0 || col >= len(mtx) {
		return newErrorFormulaArg(formulaErrorNA, "VLOOKUP has invalid column index")
	}
	if wasExact || !exactMatch {
		return mtx[col]
	}
	return newErrorFormulaArg(formulaErrorNA, "VLOOKUP no result found")
}

// vlookupBinarySearch finds the position of a target value when range lookup
// is TRUE, if the data of table array can't guarantee be sorted, it will
// return wrong result.
func vlookupBinarySearch(tableArray, lookupValue formulaArg) (matchIdx int, wasExact bool) {
	var low, high, lastMatchIdx int = 0, len(tableArray.Matrix) - 1, -1
	for low <= high {
		var mid int = low + (high-low)/2
		mtx := tableArray.Matrix[mid]
		lhs := mtx[0]
		switch lookupValue.Type {
		case ArgNumber:
			if !lookupValue.Boolean {
				lhs = mtx[0].ToNumber()
				if lhs.Type == ArgError {
					lhs = mtx[0]
				}
			}
		case ArgMatrix:
			lhs = tableArray
		}
		result := compareFormulaArg(lhs, lookupValue, false, false)
		if result == criteriaEq {
			matchIdx, wasExact = mid, true
			return
		} else if result == criteriaG {
			high = mid - 1
		} else if result == criteriaL {
			matchIdx, low = mid, mid+1
			if lhs.Value() != "" {
				lastMatchIdx = matchIdx
			}
		} else {
			return -1, false
		}
	}
	matchIdx, wasExact = lastMatchIdx, true
	return
}

// vlookupBinarySearch finds the position of a target value when range lookup
// is TRUE, if the data of table array can't guarantee be sorted, it will
// return wrong result.
func hlookupBinarySearch(row []formulaArg, lookupValue formulaArg) (matchIdx int, wasExact bool) {
	var low, high, lastMatchIdx int = 0, len(row) - 1, -1
	for low <= high {
		var mid int = low + (high-low)/2
		mtx := row[mid]
		result := compareFormulaArg(mtx, lookupValue, false, false)
		if result == criteriaEq {
			matchIdx, wasExact = mid, true
			return
		} else if result == criteriaG {
			high = mid - 1
		} else if result == criteriaL {
			low, lastMatchIdx = mid+1, mid
		} else {
			return -1, false
		}
	}
	matchIdx, wasExact = lastMatchIdx, true
	return
}

// LOOKUP function performs an approximate match lookup in a one-column or
// one-row range, and returns the corresponding value from another one-column
// or one-row range. The syntax of the function is:
//
//    LOOKUP(lookup_value,lookup_vector,[result_vector])
//
func (fn *formulaFuncs) LOOKUP(argsList *list.List) formulaArg {
	if argsList.Len() < 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "LOOKUP requires at least 2 arguments")
	}
	if argsList.Len() > 3 {
		return newErrorFormulaArg(formulaErrorVALUE, "LOOKUP requires at most 3 arguments")
	}
	lookupValue := argsList.Front().Value.(formulaArg)
	lookupVector := argsList.Front().Next().Value.(formulaArg)
	if lookupVector.Type != ArgMatrix && lookupVector.Type != ArgList {
		return newErrorFormulaArg(formulaErrorVALUE, "LOOKUP requires second argument of table array")
	}
	cols, matchIdx := lookupCol(lookupVector), -1
	for idx, col := range cols {
		lhs := lookupValue
		switch col.Type {
		case ArgNumber:
			lhs = lhs.ToNumber()
			if !col.Boolean {
				if lhs.Type == ArgError {
					lhs = lookupValue
				}
			}
		}
		if compareFormulaArg(lhs, col, false, false) == criteriaEq {
			matchIdx = idx
			break
		}
	}
	column := cols
	if argsList.Len() == 3 {
		column = lookupCol(argsList.Back().Value.(formulaArg))
	}
	if matchIdx < 0 || matchIdx >= len(column) {
		return newErrorFormulaArg(formulaErrorNA, "LOOKUP no result found")
	}
	return column[matchIdx]
}

// lookupCol extract columns for LOOKUP.
func lookupCol(arr formulaArg) []formulaArg {
	col := arr.List
	if arr.Type == ArgMatrix {
		col = nil
		for _, r := range arr.Matrix {
			if len(r) > 0 {
				col = append(col, r[0])
				continue
			}
			col = append(col, newEmptyFormulaArg())
		}
	}
	return col
}

// ROW function returns the first row number within a supplied reference or
// the number of the current row. The syntax of the function is:
//
//    ROW([reference])
//
func (fn *formulaFuncs) ROW(argsList *list.List) formulaArg {
	if argsList.Len() > 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "ROW requires at most 1 argument")
	}
	if argsList.Len() == 1 {
		if argsList.Front().Value.(formulaArg).cellRanges != nil && argsList.Front().Value.(formulaArg).cellRanges.Len() > 0 {
			return newNumberFormulaArg(float64(argsList.Front().Value.(formulaArg).cellRanges.Front().Value.(cellRange).From.Row))
		}
		if argsList.Front().Value.(formulaArg).cellRefs != nil && argsList.Front().Value.(formulaArg).cellRefs.Len() > 0 {
			return newNumberFormulaArg(float64(argsList.Front().Value.(formulaArg).cellRefs.Front().Value.(cellRef).Row))
		}
		return newErrorFormulaArg(formulaErrorVALUE, "invalid reference")
	}
	_, row, _ := CellNameToCoordinates(fn.cell)
	return newNumberFormulaArg(float64(row))
}

// ROWS function takes an Excel range and returns the number of rows that are
// contained within the range. The syntax of the function is:
//
//    ROWS(array)
//
func (fn *formulaFuncs) ROWS(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "ROWS requires 1 argument")
	}
	var min, max int
	if argsList.Front().Value.(formulaArg).cellRanges != nil && argsList.Front().Value.(formulaArg).cellRanges.Len() > 0 {
		crs := argsList.Front().Value.(formulaArg).cellRanges
		for cr := crs.Front(); cr != nil; cr = cr.Next() {
			if min == 0 {
				min = cr.Value.(cellRange).From.Row
			}
			if min > cr.Value.(cellRange).From.Row {
				min = cr.Value.(cellRange).From.Row
			}
			if min > cr.Value.(cellRange).To.Row {
				min = cr.Value.(cellRange).To.Row
			}
			if max < cr.Value.(cellRange).To.Row {
				max = cr.Value.(cellRange).To.Row
			}
			if max < cr.Value.(cellRange).From.Row {
				max = cr.Value.(cellRange).From.Row
			}
		}
	}
	if argsList.Front().Value.(formulaArg).cellRefs != nil && argsList.Front().Value.(formulaArg).cellRefs.Len() > 0 {
		cr := argsList.Front().Value.(formulaArg).cellRefs
		for refs := cr.Front(); refs != nil; refs = refs.Next() {
			if min == 0 {
				min = refs.Value.(cellRef).Row
			}
			if min > refs.Value.(cellRef).Row {
				min = refs.Value.(cellRef).Row
			}
			if max < refs.Value.(cellRef).Row {
				max = refs.Value.(cellRef).Row
			}
		}
	}
	if max == TotalRows {
		return newStringFormulaArg(strconv.Itoa(TotalRows))
	}
	result := max - min + 1
	if max == min {
		if min == 0 {
			return newErrorFormulaArg(formulaErrorVALUE, "invalid reference")
		}
		return newNumberFormulaArg(float64(1))
	}
	return newStringFormulaArg(strconv.Itoa(result))
}

// Web Functions

// ENCODEURL function returns a URL-encoded string, replacing certain
// non-alphanumeric characters with the percentage symbol (%) and a
// hexadecimal number. The syntax of the function is:
//
//    ENCODEURL(url)
//
func (fn *formulaFuncs) ENCODEURL(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "ENCODEURL requires 1 argument")
	}
	token := argsList.Front().Value.(formulaArg).Value()
	return newStringFormulaArg(strings.Replace(url.QueryEscape(token), "+", "%20", -1))
}

// Financial Functions

// CUMIPMT function calculates the cumulative interest paid on a loan or
// investment, between two specified periods. The syntax of the function is:
//
//    CUMIPMT(rate,nper,pv,start_period,end_period,type)
//
func (fn *formulaFuncs) CUMIPMT(argsList *list.List) formulaArg {
	return fn.cumip("CUMIPMT", argsList)
}

// CUMPRINC function calculates the cumulative payment on the principal of a
// loan or investment, between two specified periods. The syntax of the
// function is:
//
//    CUMPRINC(rate,nper,pv,start_period,end_period,type)
//
func (fn *formulaFuncs) CUMPRINC(argsList *list.List) formulaArg {
	return fn.cumip("CUMPRINC", argsList)
}

// cumip is an implementation of the formula function CUMIPMT and CUMPRINC.
func (fn *formulaFuncs) cumip(name string, argsList *list.List) formulaArg {
	if argsList.Len() != 6 {
		return newErrorFormulaArg(formulaErrorVALUE, fmt.Sprintf("%s requires 6 arguments", name))
	}
	rate := argsList.Front().Value.(formulaArg).ToNumber()
	if rate.Type != ArgNumber {
		return rate
	}
	nper := argsList.Front().Next().Value.(formulaArg).ToNumber()
	if nper.Type != ArgNumber {
		return nper
	}
	pv := argsList.Front().Next().Next().Value.(formulaArg).ToNumber()
	if pv.Type != ArgNumber {
		return pv
	}
	start := argsList.Back().Prev().Prev().Value.(formulaArg).ToNumber()
	if start.Type != ArgNumber {
		return start
	}
	end := argsList.Back().Prev().Value.(formulaArg).ToNumber()
	if end.Type != ArgNumber {
		return end
	}
	typ := argsList.Back().Value.(formulaArg).ToNumber()
	if typ.Type != ArgNumber {
		return typ
	}
	if typ.Number != 0 && typ.Number != 1 {
		return newErrorFormulaArg(formulaErrorNA, formulaErrorNA)
	}
	if start.Number < 1 || start.Number > end.Number {
		return newErrorFormulaArg(formulaErrorNA, formulaErrorNA)
	}
	num := 0.0
	for per := start.Number; per <= end.Number; per++ {
		args := list.New().Init()
		args.PushBack(rate)
		args.PushBack(newNumberFormulaArg(per))
		args.PushBack(nper)
		args.PushBack(pv)
		args.PushBack(newNumberFormulaArg(0))
		args.PushBack(typ)
		if name == "CUMIPMT" {
			num += fn.IPMT(args).Number
			continue
		}
		num += fn.PPMT(args).Number
	}
	return newNumberFormulaArg(num)
}

// DB function calculates the depreciation of an asset, using the Fixed
// Declining Balance Method, for each period of the asset's lifetime. The
// syntax of the function is:
//
//    DB(cost,salvage,life,period,[month])
//
func (fn *formulaFuncs) DB(argsList *list.List) formulaArg {
	if argsList.Len() < 4 {
		return newErrorFormulaArg(formulaErrorVALUE, "DB requires at least 4 arguments")
	}
	if argsList.Len() > 5 {
		return newErrorFormulaArg(formulaErrorVALUE, "DB allows at most 5 arguments")
	}
	cost := argsList.Front().Value.(formulaArg).ToNumber()
	if cost.Type != ArgNumber {
		return cost
	}
	salvage := argsList.Front().Next().Value.(formulaArg).ToNumber()
	if salvage.Type != ArgNumber {
		return salvage
	}
	life := argsList.Front().Next().Next().Value.(formulaArg).ToNumber()
	if life.Type != ArgNumber {
		return life
	}
	period := argsList.Front().Next().Next().Next().Value.(formulaArg).ToNumber()
	if period.Type != ArgNumber {
		return period
	}
	month := newNumberFormulaArg(12)
	if argsList.Len() == 5 {
		if month = argsList.Back().Value.(formulaArg).ToNumber(); month.Type != ArgNumber {
			return month
		}
	}
	if cost.Number == 0 {
		return newNumberFormulaArg(0)
	}
	if (cost.Number <= 0) || ((salvage.Number / cost.Number) < 0) || (life.Number <= 0) || (period.Number < 1) || (month.Number < 1) {
		return newErrorFormulaArg(formulaErrorNA, formulaErrorNA)
	}
	dr := 1 - math.Pow(salvage.Number/cost.Number, 1/life.Number)
	dr = math.Round(dr*1000) / 1000
	pd, depreciation := 0.0, 0.0
	for per := 1; per <= int(period.Number); per++ {
		if per == 1 {
			depreciation = cost.Number * dr * month.Number / 12
		} else if per == int(life.Number+1) {
			depreciation = (cost.Number - pd) * dr * (12 - month.Number) / 12
		} else {
			depreciation = (cost.Number - pd) * dr
		}
		pd += depreciation
	}
	return newNumberFormulaArg(depreciation)
}

// DDB function calculates the depreciation of an asset, using the Double
// Declining Balance Method, or another specified depreciation rate. The
// syntax of the function is:
//
//    DDB(cost,salvage,life,period,[factor])
//
func (fn *formulaFuncs) DDB(argsList *list.List) formulaArg {
	if argsList.Len() < 4 {
		return newErrorFormulaArg(formulaErrorVALUE, "DDB requires at least 4 arguments")
	}
	if argsList.Len() > 5 {
		return newErrorFormulaArg(formulaErrorVALUE, "DDB allows at most 5 arguments")
	}
	cost := argsList.Front().Value.(formulaArg).ToNumber()
	if cost.Type != ArgNumber {
		return cost
	}
	salvage := argsList.Front().Next().Value.(formulaArg).ToNumber()
	if salvage.Type != ArgNumber {
		return salvage
	}
	life := argsList.Front().Next().Next().Value.(formulaArg).ToNumber()
	if life.Type != ArgNumber {
		return life
	}
	period := argsList.Front().Next().Next().Next().Value.(formulaArg).ToNumber()
	if period.Type != ArgNumber {
		return period
	}
	factor := newNumberFormulaArg(2)
	if argsList.Len() == 5 {
		if factor = argsList.Back().Value.(formulaArg).ToNumber(); factor.Type != ArgNumber {
			return factor
		}
	}
	if cost.Number == 0 {
		return newNumberFormulaArg(0)
	}
	if (cost.Number <= 0) || ((salvage.Number / cost.Number) < 0) || (life.Number <= 0) || (period.Number < 1) || (factor.Number <= 0.0) || (period.Number > life.Number) {
		return newErrorFormulaArg(formulaErrorNA, formulaErrorNA)
	}
	pd, depreciation := 0.0, 0.0
	for per := 1; per <= int(period.Number); per++ {
		depreciation = math.Min((cost.Number-pd)*(factor.Number/life.Number), (cost.Number - salvage.Number - pd))
		pd += depreciation
	}
	return newNumberFormulaArg(depreciation)
}

// DOLLARDE function converts a dollar value in fractional notation, into a
// dollar value expressed as a decimal. The syntax of the function is:
//
//    DOLLARDE(fractional_dollar,fraction)
//
func (fn *formulaFuncs) DOLLARDE(argsList *list.List) formulaArg {
	return fn.dollar("DOLLARDE", argsList)
}

// DOLLARFR function converts a dollar value in decimal notation, into a
// dollar value that is expressed in fractional notation. The syntax of the
// function is:
//
//    DOLLARFR(decimal_dollar,fraction)
//
func (fn *formulaFuncs) DOLLARFR(argsList *list.List) formulaArg {
	return fn.dollar("DOLLARFR", argsList)
}

// dollar is an implementation of the formula function DOLLARDE and DOLLARFR.
func (fn *formulaFuncs) dollar(name string, argsList *list.List) formulaArg {
	if argsList.Len() != 2 {
		return newErrorFormulaArg(formulaErrorVALUE, fmt.Sprintf("%s requires 2 arguments", name))
	}
	dollar := argsList.Front().Value.(formulaArg).ToNumber()
	if dollar.Type != ArgNumber {
		return dollar
	}
	frac := argsList.Back().Value.(formulaArg).ToNumber()
	if frac.Type != ArgNumber {
		return frac
	}
	if frac.Number < 0 {
		return newErrorFormulaArg(formulaErrorNUM, formulaErrorNUM)
	}
	if frac.Number == 0 {
		return newErrorFormulaArg(formulaErrorDIV, formulaErrorDIV)
	}
	cents := math.Mod(dollar.Number, 1)
	if name == "DOLLARDE" {
		cents /= frac.Number
		cents *= math.Pow(10, math.Ceil(math.Log10(frac.Number)))
	} else {
		cents *= frac.Number
		cents *= math.Pow(10, -math.Ceil(math.Log10(frac.Number)))
	}
	return newNumberFormulaArg(math.Floor(dollar.Number) + cents)
}

// EFFECT function returns the effective annual interest rate for a given
// nominal interest rate and number of compounding periods per year. The
// syntax of the function is:
//
//    EFFECT(nominal_rate,npery)
//
func (fn *formulaFuncs) EFFECT(argsList *list.List) formulaArg {
	if argsList.Len() != 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "EFFECT requires 2 arguments")
	}
	rate := argsList.Front().Value.(formulaArg).ToNumber()
	if rate.Type != ArgNumber {
		return rate
	}
	npery := argsList.Back().Value.(formulaArg).ToNumber()
	if npery.Type != ArgNumber {
		return npery
	}
	if rate.Number <= 0 || npery.Number < 1 {
		return newErrorFormulaArg(formulaErrorNUM, formulaErrorNUM)
	}
	return newNumberFormulaArg(math.Pow((1+rate.Number/npery.Number), npery.Number) - 1)
}

// FV function calculates the Future Value of an investment with periodic
// constant payments and a constant interest rate. The syntax of the function
// is:
//
//    FV(rate,nper,[pmt],[pv],[type])
//
func (fn *formulaFuncs) FV(argsList *list.List) formulaArg {
	if argsList.Len() < 3 {
		return newErrorFormulaArg(formulaErrorVALUE, "FV requires at least 3 arguments")
	}
	if argsList.Len() > 5 {
		return newErrorFormulaArg(formulaErrorVALUE, "FV allows at most 5 arguments")
	}
	rate := argsList.Front().Value.(formulaArg).ToNumber()
	if rate.Type != ArgNumber {
		return rate
	}
	nper := argsList.Front().Next().Value.(formulaArg).ToNumber()
	if nper.Type != ArgNumber {
		return nper
	}
	pmt := argsList.Front().Next().Next().Value.(formulaArg).ToNumber()
	if pmt.Type != ArgNumber {
		return pmt
	}
	pv, typ := newNumberFormulaArg(0), newNumberFormulaArg(0)
	if argsList.Len() >= 4 {
		if pv = argsList.Front().Next().Next().Next().Value.(formulaArg).ToNumber(); pv.Type != ArgNumber {
			return pv
		}
	}
	if argsList.Len() == 5 {
		if typ = argsList.Back().Value.(formulaArg).ToNumber(); typ.Type != ArgNumber {
			return typ
		}
	}
	if typ.Number != 0 && typ.Number != 1 {
		return newErrorFormulaArg(formulaErrorNA, formulaErrorNA)
	}
	if rate.Number != 0 {
		return newNumberFormulaArg(-pv.Number*math.Pow(1+rate.Number, nper.Number) - pmt.Number*(1+rate.Number*typ.Number)*(math.Pow(1+rate.Number, nper.Number)-1)/rate.Number)
	}
	return newNumberFormulaArg(-pv.Number - pmt.Number*nper.Number)
}

// FVSCHEDULE function calculates the Future Value of an investment with a
// variable interest rate. The syntax of the function is:
//
//    FVSCHEDULE(principal,schedule)
//
func (fn *formulaFuncs) FVSCHEDULE(argsList *list.List) formulaArg {
	if argsList.Len() != 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "FVSCHEDULE requires 2 arguments")
	}
	pri := argsList.Front().Value.(formulaArg).ToNumber()
	if pri.Type != ArgNumber {
		return pri
	}
	principal := pri.Number
	for _, arg := range argsList.Back().Value.(formulaArg).ToList() {
		if arg.Value() == "" {
			continue
		}
		rate := arg.ToNumber()
		if rate.Type != ArgNumber {
			return rate
		}
		principal *= (1 + rate.Number)
	}
	return newNumberFormulaArg(principal)
}

// IPMT function calculates the interest payment, during a specific period of a
// loan or investment that is paid in constant periodic payments, with a
// constant interest rate. The syntax of the function is:
//
//    IPMT(rate,per,nper,pv,[fv],[type])
//
func (fn *formulaFuncs) IPMT(argsList *list.List) formulaArg {
	return fn.ipmt("IPMT", argsList)
}

// ipmt is an implementation of the formula function IPMT and PPMT.
func (fn *formulaFuncs) ipmt(name string, argsList *list.List) formulaArg {
	if argsList.Len() < 4 {
		return newErrorFormulaArg(formulaErrorVALUE, fmt.Sprintf("%s requires at least 4 arguments", name))
	}
	if argsList.Len() > 6 {
		return newErrorFormulaArg(formulaErrorVALUE, fmt.Sprintf("%s allows at most 6 arguments", name))
	}
	rate := argsList.Front().Value.(formulaArg).ToNumber()
	if rate.Type != ArgNumber {
		return rate
	}
	per := argsList.Front().Next().Value.(formulaArg).ToNumber()
	if per.Type != ArgNumber {
		return per
	}
	nper := argsList.Front().Next().Next().Value.(formulaArg).ToNumber()
	if nper.Type != ArgNumber {
		return nper
	}
	pv := argsList.Front().Next().Next().Next().Value.(formulaArg).ToNumber()
	if pv.Type != ArgNumber {
		return pv
	}
	fv, typ := newNumberFormulaArg(0), newNumberFormulaArg(0)
	if argsList.Len() >= 5 {
		if fv = argsList.Front().Next().Next().Next().Next().Value.(formulaArg).ToNumber(); fv.Type != ArgNumber {
			return fv
		}
	}
	if argsList.Len() == 6 {
		if typ = argsList.Back().Value.(formulaArg).ToNumber(); typ.Type != ArgNumber {
			return typ
		}
	}
	if typ.Number != 0 && typ.Number != 1 {
		return newErrorFormulaArg(formulaErrorNA, formulaErrorNA)
	}
	if per.Number <= 0 || per.Number > nper.Number {
		return newErrorFormulaArg(formulaErrorNA, formulaErrorNA)
	}
	args := list.New().Init()
	args.PushBack(rate)
	args.PushBack(nper)
	args.PushBack(pv)
	args.PushBack(fv)
	args.PushBack(typ)
	pmt, capital, interest, principal := fn.PMT(args), pv.Number, 0.0, 0.0
	for i := 1; i <= int(per.Number); i++ {
		if typ.Number != 0 && i == 1 {
			interest = 0
		} else {
			interest = -capital * rate.Number
		}
		principal = pmt.Number - interest
		capital += principal
	}
	if name == "IPMT" {
		return newNumberFormulaArg(interest)
	}
	return newNumberFormulaArg(principal)
}

// IRR function returns the Internal Rate of Return for a supplied series of
// periodic cash flows (i.e. an initial investment value and a series of net
// income values). The syntax of the function is:
//
//    IRR(values,[guess])
//
func (fn *formulaFuncs) IRR(argsList *list.List) formulaArg {
	if argsList.Len() < 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "IRR requires at least 1 argument")
	}
	if argsList.Len() > 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "IRR allows at most 2 arguments")
	}
	values, guess := argsList.Front().Value.(formulaArg).ToList(), newNumberFormulaArg(0.1)
	if argsList.Len() > 1 {
		if guess = argsList.Back().Value.(formulaArg).ToNumber(); guess.Type != ArgNumber {
			return guess
		}
	}
	x1, x2 := newNumberFormulaArg(0), guess
	args := list.New().Init()
	args.PushBack(x1)
	for _, v := range values {
		args.PushBack(v)
	}
	f1 := fn.NPV(args)
	args.Front().Value = x2
	f2 := fn.NPV(args)
	for i := 0; i < maxFinancialIterations; i++ {
		if f1.Number*f2.Number < 0 {
			break
		}
		if math.Abs(f1.Number) < math.Abs((f2.Number)) {
			x1.Number += 1.6 * (x1.Number - x2.Number)
			args.Front().Value = x1
			f1 = fn.NPV(args)
			continue
		}
		x2.Number += 1.6 * (x2.Number - x1.Number)
		args.Front().Value = x2
		f2 = fn.NPV(args)
	}
	if f1.Number*f2.Number > 0 {
		return newErrorFormulaArg(formulaErrorNUM, formulaErrorNUM)
	}
	args.Front().Value = x1
	f := fn.NPV(args)
	var rtb, dx, xMid, fMid float64
	if f.Number < 0 {
		rtb = x1.Number
		dx = x2.Number - x1.Number
	} else {
		rtb = x2.Number
		dx = x1.Number - x2.Number
	}
	for i := 0; i < maxFinancialIterations; i++ {
		dx *= 0.5
		xMid = rtb + dx
		args.Front().Value = newNumberFormulaArg(xMid)
		fMid = fn.NPV(args).Number
		if fMid <= 0 {
			rtb = xMid
		}
		if math.Abs(fMid) < financialPercision || math.Abs(dx) < financialPercision {
			break
		}
	}
	return newNumberFormulaArg(xMid)
}

// ISPMT function calculates the interest paid during a specific period of a
// loan or investment. The syntax of the function is:
//
//    ISPMT(rate,per,nper,pv)
//
func (fn *formulaFuncs) ISPMT(argsList *list.List) formulaArg {
	if argsList.Len() != 4 {
		return newErrorFormulaArg(formulaErrorVALUE, "ISPMT requires 4 arguments")
	}
	rate := argsList.Front().Value.(formulaArg).ToNumber()
	if rate.Type != ArgNumber {
		return rate
	}
	per := argsList.Front().Next().Value.(formulaArg).ToNumber()
	if per.Type != ArgNumber {
		return per
	}
	nper := argsList.Back().Prev().Value.(formulaArg).ToNumber()
	if nper.Type != ArgNumber {
		return nper
	}
	pv := argsList.Back().Value.(formulaArg).ToNumber()
	if pv.Type != ArgNumber {
		return pv
	}
	pr, payment, num := pv.Number, pv.Number/nper.Number, 0.0
	for i := 0; i <= int(per.Number); i++ {
		num = rate.Number * pr * -1
		pr -= payment
		if i == int(nper.Number) {
			num = 0
		}
	}
	return newNumberFormulaArg(num)
}

// MIRR function returns the Modified Internal Rate of Return for a supplied
// series of periodic cash flows (i.e. a set of values, which includes an
// initial investment value and a series of net income values). The syntax of
// the function is:
//
//    MIRR(values,finance_rate,reinvest_rate)
//
func (fn *formulaFuncs) MIRR(argsList *list.List) formulaArg {
	if argsList.Len() != 3 {
		return newErrorFormulaArg(formulaErrorVALUE, "MIRR requires 3 arguments")
	}
	values := argsList.Front().Value.(formulaArg).ToList()
	financeRate := argsList.Front().Next().Value.(formulaArg).ToNumber()
	if financeRate.Type != ArgNumber {
		return financeRate
	}
	reinvestRate := argsList.Back().Value.(formulaArg).ToNumber()
	if reinvestRate.Type != ArgNumber {
		return reinvestRate
	}
	n, fr, rr, npvPos, npvNeg := len(values), 1+financeRate.Number, 1+reinvestRate.Number, 0.0, 0.0
	for i, v := range values {
		val := v.ToNumber()
		if val.Number >= 0 {
			npvPos += val.Number / math.Pow(float64(rr), float64(i))
			continue
		}
		npvNeg += val.Number / math.Pow(float64(fr), float64(i))
	}
	if npvNeg == 0 || npvPos == 0 || reinvestRate.Number <= -1 {
		return newErrorFormulaArg(formulaErrorDIV, formulaErrorDIV)
	}
	return newNumberFormulaArg(math.Pow(-npvPos*math.Pow(rr, float64(n))/(npvNeg*rr), 1/(float64(n)-1)) - 1)
}

// NOMINAL function returns the nominal interest rate for a given effective
// interest rate and number of compounding periods per year. The syntax of
// the function is:
//
//    NOMINAL(effect_rate,npery)
//
func (fn *formulaFuncs) NOMINAL(argsList *list.List) formulaArg {
	if argsList.Len() != 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "NOMINAL requires 2 arguments")
	}
	rate := argsList.Front().Value.(formulaArg).ToNumber()
	if rate.Type != ArgNumber {
		return rate
	}
	npery := argsList.Back().Value.(formulaArg).ToNumber()
	if npery.Type != ArgNumber {
		return npery
	}
	if rate.Number <= 0 || npery.Number < 1 {
		return newErrorFormulaArg(formulaErrorNUM, formulaErrorNUM)
	}
	return newNumberFormulaArg(npery.Number * (math.Pow(rate.Number+1, 1/npery.Number) - 1))
}

// NPER function calculates the number of periods required to pay off a loan,
// for a constant periodic payment and a constant interest rate. The syntax
// of the function is:
//
//    NPER(rate,pmt,pv,[fv],[type])
//
func (fn *formulaFuncs) NPER(argsList *list.List) formulaArg {
	if argsList.Len() < 3 {
		return newErrorFormulaArg(formulaErrorVALUE, "NPER requires at least 3 arguments")
	}
	if argsList.Len() > 5 {
		return newErrorFormulaArg(formulaErrorVALUE, "NPER allows at most 5 arguments")
	}
	rate := argsList.Front().Value.(formulaArg).ToNumber()
	if rate.Type != ArgNumber {
		return rate
	}
	pmt := argsList.Front().Next().Value.(formulaArg).ToNumber()
	if pmt.Type != ArgNumber {
		return pmt
	}
	pv := argsList.Front().Next().Next().Value.(formulaArg).ToNumber()
	if pv.Type != ArgNumber {
		return pv
	}
	fv, typ := newNumberFormulaArg(0), newNumberFormulaArg(0)
	if argsList.Len() >= 4 {
		if fv = argsList.Front().Next().Next().Next().Value.(formulaArg).ToNumber(); fv.Type != ArgNumber {
			return fv
		}
	}
	if argsList.Len() == 5 {
		if typ = argsList.Back().Value.(formulaArg).ToNumber(); typ.Type != ArgNumber {
			return typ
		}
	}
	if typ.Number != 0 && typ.Number != 1 {
		return newErrorFormulaArg(formulaErrorNA, formulaErrorNA)
	}
	if pmt.Number == 0 {
		return newErrorFormulaArg(formulaErrorNUM, formulaErrorNUM)
	}
	if rate.Number != 0 {
		p := math.Log((pmt.Number*(1+rate.Number*typ.Number)/rate.Number-fv.Number)/(pv.Number+pmt.Number*(1+rate.Number*typ.Number)/rate.Number)) / math.Log(1+rate.Number)
		return newNumberFormulaArg(p)
	}
	return newNumberFormulaArg((-pv.Number - fv.Number) / pmt.Number)
}

// NPV function calculates the Net Present Value of an investment, based on a
// supplied discount rate, and a series of future payments and income. The
// syntax of the function is:
//
//    NPV(rate,value1,[value2],[value3],...)
//
func (fn *formulaFuncs) NPV(argsList *list.List) formulaArg {
	if argsList.Len() < 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "NPV requires at least 2 arguments")
	}
	rate := argsList.Front().Value.(formulaArg).ToNumber()
	if rate.Type != ArgNumber {
		return rate
	}
	val, i := 0.0, 1
	for arg := argsList.Front().Next(); arg != nil; arg = arg.Next() {
		num := arg.Value.(formulaArg).ToNumber()
		if num.Type != ArgNumber {
			continue
		}
		val += num.Number / math.Pow(1+rate.Number, float64(i))
		i++
	}
	return newNumberFormulaArg(val)
}

// PDURATION function calculates the number of periods required for an
// investment to reach a specified future value. The syntax of the function
// is:
//
//    PDURATION(rate,pv,fv)
//
func (fn *formulaFuncs) PDURATION(argsList *list.List) formulaArg {
	if argsList.Len() != 3 {
		return newErrorFormulaArg(formulaErrorVALUE, "PDURATION requires 3 arguments")
	}
	rate := argsList.Front().Value.(formulaArg).ToNumber()
	if rate.Type != ArgNumber {
		return rate
	}
	pv := argsList.Front().Next().Value.(formulaArg).ToNumber()
	if pv.Type != ArgNumber {
		return pv
	}
	fv := argsList.Back().Value.(formulaArg).ToNumber()
	if fv.Type != ArgNumber {
		return fv
	}
	if rate.Number <= 0 || pv.Number <= 0 || fv.Number <= 0 {
		return newErrorFormulaArg(formulaErrorNUM, formulaErrorNUM)
	}
	return newNumberFormulaArg((math.Log(fv.Number) - math.Log(pv.Number)) / math.Log(1+rate.Number))
}

// PMT function calculates the constant periodic payment required to pay off
// (or partially pay off) a loan or investment, with a constant interest
// rate, over a specified period. The syntax of the function is:
//
//    PMT(rate,nper,pv,[fv],[type])
//
func (fn *formulaFuncs) PMT(argsList *list.List) formulaArg {
	if argsList.Len() < 3 {
		return newErrorFormulaArg(formulaErrorVALUE, "PMT requires at least 3 arguments")
	}
	if argsList.Len() > 5 {
		return newErrorFormulaArg(formulaErrorVALUE, "PMT allows at most 5 arguments")
	}
	rate := argsList.Front().Value.(formulaArg).ToNumber()
	if rate.Type != ArgNumber {
		return rate
	}
	nper := argsList.Front().Next().Value.(formulaArg).ToNumber()
	if nper.Type != ArgNumber {
		return nper
	}
	pv := argsList.Front().Next().Next().Value.(formulaArg).ToNumber()
	if pv.Type != ArgNumber {
		return pv
	}
	fv, typ := newNumberFormulaArg(0), newNumberFormulaArg(0)
	if argsList.Len() >= 4 {
		if fv = argsList.Front().Next().Next().Next().Value.(formulaArg).ToNumber(); fv.Type != ArgNumber {
			return fv
		}
	}
	if argsList.Len() == 5 {
		if typ = argsList.Back().Value.(formulaArg).ToNumber(); typ.Type != ArgNumber {
			return typ
		}
	}
	if typ.Number != 0 && typ.Number != 1 {
		return newErrorFormulaArg(formulaErrorNA, formulaErrorNA)
	}
	if rate.Number != 0 {
		p := (-fv.Number - pv.Number*math.Pow((1+rate.Number), nper.Number)) / (1 + rate.Number*typ.Number) / ((math.Pow((1+rate.Number), nper.Number) - 1) / rate.Number)
		return newNumberFormulaArg(p)
	}
	return newNumberFormulaArg((-pv.Number - fv.Number) / nper.Number)
}

// PPMT function calculates the payment on the principal, during a specific
// period of a loan or investment that is paid in constant periodic payments,
// with a constant interest rate. The syntax of the function is:
//
//    PPMT(rate,per,nper,pv,[fv],[type])
//
func (fn *formulaFuncs) PPMT(argsList *list.List) formulaArg {
	return fn.ipmt("PPMT", argsList)
}
