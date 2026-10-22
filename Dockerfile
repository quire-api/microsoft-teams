# Pin the Node patch version: a floating `node:22-alpine` leaves the Node that
# ships up to whatever the builder has cached. Bump on each Node security
# release (zkoss/boeneo#26073).
FROM node:22.23.4-alpine
ARG TARGETARCH
RUN sed -i 's/dl-cdn.alpinelinux.org/mirror.twds.com.tw/g' /etc/apk/repositories \
&& apk add --update --no-cache aws-cli
COPY ./ /app
RUN cd /app \
&& npm ci --omit=dev
ENTRYPOINT [ "/bin/sh", "/app/start_app.sh" ]
EXPOSE 3978
