package kafka

import (
	"context"

	"github.com/Shopify/sarama"
)

// Handler message from mq.
type Handler func(ctx context.Context, message []byte)

// Publish message to mq.
type Publish func(message []byte) error

type handler struct {
	partitionConsumer sarama.PartitionConsumer
	handler           Handler
}
