package main

import (
	"context"
	"encoding/json"
	"fmt"
	"os"
	"os/signal"
	"syscall"
	"time"

	"github.com/go-kit/kit/log"
	"github.com/go-kit/kit/log/level"
	"github.com/kelseyhightower/envconfig"

	"github.com/geoirb/go-templater/internal/kafka"
	"github.com/geoirb/go-templater/internal/parser"
	"github.com/geoirb/go-templater/internal/path"
	"github.com/geoirb/go-templater/internal/placeholder"
	"github.com/geoirb/go-templater/internal/qrcode"
	"github.com/geoirb/go-templater/internal/response"
	"github.com/geoirb/go-templater/internal/templater"
	"github.com/geoirb/go-templater/internal/templater/mq"
	"github.com/geoirb/go-templater/internal/xlsx"
)

type configuration struct {
	MQHost string `envconfig:"MQ_HOST" default:"localhost"`
	MQPort int    `envconfig:"MQ_PORT" default:"9093"`

	TemplateDir string `envconfig:"TEMPLATE_DIR" default:"/template"`

	FillInTopicRequest  string `envconfig:"FILL_IN_TOPIC_REQUEST" default:"request"`
	FillInTopicResponse string `envconfig:"FILL_IN_TOPIC_RESPONSE" default:"response"`
}

const (
	prefixCfg   = ""
	serviceName = "templater"
)

func main() {
	fmt.Println("trial version")
	if time.Since(time.Date(2021, time.September, 13, 0, 0, 0, 0, time.Now().Location())) > 0 {
		return
	}

	logger := log.NewJSONLogger(log.NewSyncWriter(os.Stdout))
	logger = log.WithPrefix(logger, "service", serviceName)
	logger = log.With(logger, "ts", log.DefaultTimestampUTC)

	var cfg configuration
	if err := envconfig.Process(prefixCfg, &cfg); err != nil {
		level.Error(logger).Log("msg", "configuration", "err", err)
		os.Exit(1)
	}

	level.Info(logger).Log("msg", "initialization", "err")

	path, err := path.NewBuilder(
		cfg.TemplateDir,
	)

	if err != nil {
		level.Error(logger).Log("msg", "path init", "err", err)
		os.Exit(1)
	}

	parser, err := parser.New()
	if err != nil {
		level.Error(logger).Log("msg", "parser init", "err", err)
		os.Exit(1)
	}

	placeholder, err := placeholder.New()
	if err != nil {
		level.Error(logger).Log("msg", "placeholder init", "err", err)
		os.Exit(1)
	}

	qrcode := qrcode.NewCreator()

	x := xlsx.NewFacade(
		placeholder,
		qrcode,
	)

	data, _ := os.ReadFile(os.Args[2])
	var payload interface{}
	json.Unmarshal(data, &payload)
	start := time.Now()
	x.FillIn(
		context.Background(),
		os.Args[1],
		payload,
	)
	fmt.Println(time.Since(start).Seconds())

	svc := templater.NewService(
		path,
		parser,
		x.FillIn,
		logger,
	)

	address := fmt.Sprintf("%s:%d", cfg.MQHost, cfg.MQPort)
	mqKafka, err := kafka.NewMessageQueue(
		[]string{address},
	)
	if err != nil {
		level.Error(logger).Log("msg", "kafka init", "address", address, "err", err)
		os.Exit(1)
	}

	handler := mq.NewFillInHandler(
		svc,
		mq.NewFillInTransport(
			response.Build,
		),
		mqKafka.NewPublish(cfg.FillInTopicResponse),
	)

	mqKafka.Consume(cfg.FillInTopicRequest, handler)

	go func() {
		level.Info(logger).Log("msg", "kafka listener turn on")
		mqKafka.ListenAndServe()
	}()

	c := make(chan os.Signal, 1)
	signal.Notify(c, syscall.SIGTERM, syscall.SIGINT)
	level.Info(logger).Log("msg", "received signal", "signal", <-c)

	level.Info(logger).Log("msg", "kafka listener shutdown")
	mqKafka.Shutdown()
	level.Info(logger).Log("msg", "stop service")
}
