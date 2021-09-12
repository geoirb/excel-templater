package templater

// Info for fill in template.
type Request struct {
	UUID     string
	UserID   int
	Template string
	Payload  interface{}
}

type Response struct {
	UUID     string
	UserID   int
	Document []byte
	Message  string
}
