FROM node:16.20.2-alpine
ARG TARGETARCH
RUN sed -i 's/dl-cdn.alpinelinux.org/mirror.twds.com.tw/g' /etc/apk/repositories \
&& apk add --update --no-cache aws-cli
COPY ./ /app
RUN cd /app \
&& npm update \
&& npm install
ENTRYPOINT [ "/bin/sh", "/app/start_app.sh" ]
EXPOSE 3978
