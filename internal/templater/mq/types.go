package mq

type request struct {
	UserID   int         `json:"user_id"`
	UUID     string      `json:"uuid"`
	Template string      `json:"template"`
	Payload  interface{} `json:"payload"`
}

type response struct {
	IsSuccess bool   `json:"is_success"`
	UserID    int    `json:"user_id"`
	UUID      string `json:"uuid"`
	Document  []byte `json:"document,omitempty"`
	Error     string `json:"error,omitempty"`
}
