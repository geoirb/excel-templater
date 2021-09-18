package xlsx

import (
	qr "github.com/skip2/go-qrcode"
)

// Encode a QR Code and return a raw PNG image.
func Encode(payload string, size int) ([]byte, error) {
	return qr.Encode(payload, qr.Medium, size)
}
