package response

import (
	"encoding/json"
)

// Build response to payload and error.
func Build(payload interface{}, err error) ([]byte, error) {
	response := response{
		IsOk: err == nil,
	}

	if payload != nil {
		response.Payload = payload
	}

	if !response.IsOk {
		response.Payload = err.Error()
	}
	return json.Marshal(response)
}
