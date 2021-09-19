package xlsx

import (
	qr "github.com/skip2/go-qrcode"
)

func encode(payload string, size int) ([]byte, error) {
	return qr.Encode(payload, qr.Medium, size)
}
