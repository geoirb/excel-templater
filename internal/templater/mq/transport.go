package mq

import (
	"encoding/json"

	"github.com/geoirb/go-templater/internal/templater"
)

// FillInTransport ...
type FillInTransport struct {
}

func NewFillInTransport() *FillInTransport {
	return &FillInTransport{}
}

// DecodeRequest ...
func (t *FillInTransport) DecodeRequest(message []byte) (templater.Request, error) {
	var req request
	err := json.Unmarshal(message, &req)
	return templater.Request(req), err
}

// DecodeRequest ...
func (t *FillInTransport) EncodeResponse(res templater.Response) (message []byte) {
	payload := response{
		UUID:      res.UUID,
		UserID:    res.UserID,
		IsSuccess: len(res.Error) == 0,
		Document:  res.Document,
		Error:     res.Error,
	}
	message, _ = json.Marshal(payload)
	return
}
