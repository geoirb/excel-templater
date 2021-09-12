package qrcode

import (
	qrcode "github.com/skip2/go-qrcode"
)

// Creator qr codes.
type Creator struct {
}

// NewCreator ...
func NewCreator() *Creator {
	return &Creator{}
}

// Create returns qr code bytes by error.
func (c *Creator) Create(payload string, size int) ([]byte, error) {
	return qrcode.Encode(payload, qrcode.Medium, size)
}
