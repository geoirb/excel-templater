package response

type response struct {
	IsOk    bool        `json:"is_ok"`
	Payload interface{} `json:"payload,omitempty"`
}
