package mq

type request struct {
	UUID     string      `json:"uuid"`
	UserID   int         `json:"user_id"`
	Template string      `json:"template"`
	Payload  interface{} `json:"payload"`
}

type response struct {
	UUID     string `json:"uuid"`
	UserID   int    `json:"user_id"`
	Document []byte `json:"document,omitempty"`
	Message  string `json:"message,omitempty"`
}
