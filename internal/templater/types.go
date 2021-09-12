package templater

// Info for fill in template.
type Request struct {
	UserID   int
	UUID     string
	Template string
	Payload  interface{}
}

type Response struct {
	UUID     string
	UserID   int
	Document []byte
	Error    string
}
