package response

import (
	"encoding/json"
	"errors"
	"testing"

	"github.com/stretchr/testify/assert"
)

var (
	testData = "test-data"
	errTest  = errors.New("test-error")
)

func TestBuild(t *testing.T) {
	t.Run("success payload", func(t *testing.T) {
		payload := struct {
			Data string `json:"data"`
		}{
			Data: testData,
		}

		response := response{
			IsOk:    true,
			Payload: payload,
		}
		expectedData, err := json.Marshal(&response)
		assert.NotNil(t, expectedData)
		assert.NoError(t, err)

		actualData, err := Build(payload, nil)
		assert.NotNil(t, actualData)
		assert.NoError(t, err)
		assert.Equal(t, expectedData, actualData)
	})

	t.Run("error", func(t *testing.T) {
		payload := struct {
			Data string `json:"data"`
		}{
			Data: testData,
		}

		response := response{
			IsOk:    false,
			Payload: errTest.Error(),
		}
		expectedData, err := json.Marshal(&response)
		assert.NotNil(t, expectedData)
		assert.NoError(t, err)

		actualData, err := Build(payload, errTest)
		assert.NotNil(t, actualData)
		assert.NoError(t, err)
		assert.Equal(t, expectedData, actualData)
	})
}
