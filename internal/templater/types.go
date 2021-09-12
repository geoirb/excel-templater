package templater

var (
	FieldNameType = "field_name"
	ArrayType     = "array"
	QRCodeType    = "qr_code"
)

// Info for fill in template.
type Request struct {
	UUID     string
	UserID   int
	Template string
	Payload  interface{}
}

type Response struct {
	UUID     string
	UserID   int
	Document []byte
	Message  string
}
