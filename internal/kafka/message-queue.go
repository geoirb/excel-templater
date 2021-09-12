package kafka

import (
	"context"

	"github.com/Shopify/sarama"
)

// MessageQueue of kafka.
type MessageQueue struct {
	ctx    context.Context
	cancel context.CancelFunc

	client   sarama.Client
	producer sarama.SyncProducer
	consumer sarama.Consumer
	handler  map[string]handler
}

// NewMessageQueue ...
func NewMessageQueue(
	addrs []string,
) (mq *MessageQueue, err error) {
	mq = &MessageQueue{
		handler: make(map[string]handler),
	}

	cfg := sarama.NewConfig()
	cfg.Producer.Return.Successes = true
	if mq.client, err = sarama.NewClient(addrs, cfg); err != nil {
		return
	}
	if mq.producer, err = sarama.NewSyncProducerFromClient(mq.client); err != nil {
		return
	}

	mq.consumer, err = sarama.NewConsumerFromClient(mq.client)
	return
}

// Consume adds consume topic.
func (mq *MessageQueue) Consume(topic string, h Handler) {
	if _, isExist := mq.handler[topic]; isExist {
		panic(errTopicIsExist)
	}

	cp, err := mq.consumer.ConsumePartition(topic, 0, 0)
	if err != nil {
		panic(err)
	}
	mq.handler[topic] = handler{
		partitionConsumer: cp,
		handler:           h,
	}
}

// NewPublish returns publish func.
func (mq *MessageQueue) NewPublish(topic string) Publish {
	return func(message []byte) (err error) {
		msg := &sarama.ProducerMessage{
			Topic: topic,
			Value: sarama.ByteEncoder(message),
		}
		_, _, err = mq.producer.SendMessage(msg)
		return
	}
}

// ListenAndServe message queue.
func (mq *MessageQueue) ListenAndServe() {
	mq.ctx, mq.cancel = context.WithCancel(context.Background())
	for _, topic := range mq.handler {
		go mq.runtime(topic)
	}
}

// Shutdown consumers message queue.
func (mq *MessageQueue) Shutdown() {
	mq.cancel()
	mq.consumer.Close()
	mq.producer.Close()
}

func (mq *MessageQueue) runtime(topic handler) {
	defer topic.partitionConsumer.Close()
	for {
		select {
		case <-mq.ctx.Done():
			return
		case m := <-topic.partitionConsumer.Messages():
			go topic.handler(mq.ctx, m.Value)
		}
	}
}
