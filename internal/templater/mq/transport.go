package mq

import (
	"encoding/json"

	"github.com/geoirb/go-templater/internal/templater"
)

type builder func(payload interface{}, err error) ([]byte, error)

// FillInTransport ...
type FillInTransport struct {
	builder builder
}

func NewFillInTransport(
	builder builder,
) *FillInTransport {
	return &FillInTransport{
		builder: builder,
	}
}

// DecodeRequest ...
func (t *FillInTransport) DecodeRequest(message []byte) (templater.Request, error) {
	var req request
	err := json.Unmarshal(message, &req)
	return templater.Request(req), err
}

// DecodeRequest ...
func (t *FillInTransport) EncodeResponse(res templater.Response, err error) (message []byte) {
	payload := response(res)
	message, _ = t.builder(payload, err)
	return
}
