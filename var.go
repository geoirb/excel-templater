package xlsx

const (
	placeholderGroupRegexp = "{([_a-zA-Z0-9:]+)}$"
	placeholderReqexp      = "([_a-zA-Z0-9]+)"
	arrayRegexp            = "array$"
	qrCodeRegexp           = "qr_code_[_a-zA-Z0-9]+"
	imageReqexp            = "image_[_a-zA-Z0-9]+"
)

const (
	defaultImage = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABAQMAAAAl21bKAAAAA1BMVEUAAACnej3aAAAAAXRSTlMAQObYZgAAAApJREFUCNdjYAAAAAIAAeIhvDMAAAAASUVORK5CYII="

	// Placeholder types.
	FieldNameType = "field_name"
	ArrayType     = "array"
	QRCodeType    = "qr_code"
	ImageType     = "image"
)

