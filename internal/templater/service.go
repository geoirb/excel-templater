package templater

import (
	"context"
	"os"

	"github.com/go-kit/kit/log"
	"github.com/go-kit/kit/log/level"
)

type fillInFunc func(ctx context.Context, template, tmpFile string, payload interface{}) error

type path interface {
	Template(name string) string
	TmpFile(name string) string
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

	templatePath, tmpFilePath := s.path.Template(req.Template), s.path.TmpFile(req.Template)
	if err = fillIn(ctx, templatePath, tmpFilePath, req.Payload); err != nil {
		level.Error(logger).Log("msg", "fill in template", "template", templateType, "err", err)
		return
	}

	defer os.Remove(tmpFilePath)
	if res.Document, err = os.ReadFile(tmpFilePath); err != nil {
		level.Error(logger).Log("msg", "read tmp file", "tmp file name", tmpFilePath, "err", err)
	}
	return
}
