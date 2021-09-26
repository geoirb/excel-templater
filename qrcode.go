package excel

import (
	qr "github.com/skip2/go-qrcode"
)

var (
	emptyStr = " "
)

func encode(payload string, size int) ([]byte, error) {
	if len(payload) == 0 {
		payload = emptyStr
	}
	return qr.Encode(payload, qr.Medium, size)
}
