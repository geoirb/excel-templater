package templater

import (
	"context"
	"io"
	"io/ioutil"

	"github.com/go-kit/kit/log"
	"github.com/go-kit/kit/log/level"
)

type fillInFunc func(ctx context.Context, template string, payload interface{}) (io.Reader, error)

type path interface {
	Template(name string) string
}

type parser interface {
	Type(filename string) (string, error)
}

type service struct {
	fillIn map[string]fillInFunc

	path   path
	parser parser

	logger log.Logger
}

func NewService(
	path path,
	parser parser,

	xlsx fillInFunc,

	logger log.Logger,
) Service {
	s := &service{
		path:   path,
		parser: parser,
		logger: logger,
	}

	s.fillIn = map[string]fillInFunc{
		"xlsx": xlsx,
	}
	return s
}

// FillIn fills template by req.
func (s *service) FillIn(ctx context.Context, req Request) (res Response, err error) {
	logger := log.WithPrefix(s.logger, "method", "FillIn", "uuid", req.UUID)

	templateType, err := s.parser.Type(req.Template)
	if err != nil {
		level.Error(logger).Log("msg", "template type", "err", err)
		return
	}
	fillIn, isExist := s.fillIn[templateType]
	if !isExist {
		level.Error(logger).Log("msg", "unknown type", "type", templateType)
		err = errUnknownTemplateType
		return
	}

	res = Response{
		UUID:   req.UUID,
		UserID: req.UserID,
	}

	templatePath := s.path.Template(req.Template)
	result, err := fillIn(ctx, templatePath, req.Payload)
	if err != nil {
		level.Error(logger).Log("msg", "fill in template", "template", templateType, "err", err)
		return
	}
	if res.Document, err = ioutil.ReadAll(result); err != nil {
		level.Error(logger).Log("msg", "read result", "err", err)
	}
	return
}
