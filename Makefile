BINARY_NAME 		= go-templater

IMAGE_NAME 			= $(BINARY_NAME)
PWD					= $(shell pwd)
GOPATH            	= $(HOME)/go
GOOS              	= linux
GOARCH            	= amd64


GO_BUILD_ARGS = 						\
				-o build/$(BINARY_NAME)	\
				cmd/go-templater/main.go

DOCKER_BUILD_ROOT  = /service
DOCKER_ARGS = 												\
   				-e GODEBUG=x509ignoreCN=0  					\
   				-e GOOS=$(GOOS)            					\
   				-e GOARCH=$(GOARCH) 						\
   				-e CGO_ENABLED=0 							\
				-v $(HOME)/.ssh:/root/.ssh:ro             	\
				-v $(HOME)/.gitconfig:/root/.gitconfig:ro 	\
				-v $(GOPATH):/go                          	\
				-v $(PWD):$(DOCKER_BUILD_ROOT)            	\
				--workdir $(DOCKER_BUILD_ROOT)            	\


lint:
	docker run 	--rm 						    			\
				$(DOCKER_ARGS) 								\
				golangci/golangci-lint:latest  				\
				golangci-lint run -v 

.PHONY: build
build:
	docker run --rm $(DOCKER_ARGS) golang go build $(GO_BUILD_ARGS)
	docker build -t $(IMAGE_NAME) -f docker/Dockerfile .

run: 
	@make build
	docker-compose -f "deployments/docker-compose.yml" up -d 

