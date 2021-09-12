package templater

import (
	"context"
)

// Service of templater.
type Service interface {
	FillIn(ctx context.Context, req Request) (res Response, err error)
	// TODO
	// Del before
	Xlsx(ctx context.Context, template, tmpFile string, payload interface{}) (err error)
}
