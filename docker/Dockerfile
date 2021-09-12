FROM alpine

RUN adduser --system appuser
WORKDIR /app
COPY build/go-templater /app/service
USER appuser

CMD ["/app/service"]