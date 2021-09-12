package mq

import (
	"context"

	"github.com/geoirb/go-templater/internal/kafka"
	"github.com/geoirb/go-templater/internal/templater"
)

type fillInServe struct {
	svc       templater.Service
	transport *FillInTransport
	Publish   kafka.Publish
}

func (s *fillInServe) Handle(ctx context.Context, message []byte) {
	request, err := s.transport.DecodeRequest(message)

	var res templater.Response
	if err == nil {
		res = s.svc.FillIn(ctx, request)
	}

	s.Publish(s.transport.EncodeResponse(res))
}

func NewFillInHandler(
	svc templater.Service,
	transport *FillInTransport,
	publish kafka.Publish,
) kafka.Handler {
	s := &fillInServe{
		svc:       svc,
		transport: transport,
		Publish:   publish,
	}

	return s.Handle
}
