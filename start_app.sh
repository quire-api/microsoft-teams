#!/bin/sh
get_ssm_parameter() {
    /usr/bin/aws ssm get-parameter --name "$1" --with-decryption \
    --query "Parameter.Value" --output text --no-cli-pager |
    base64 -d | tee /app/env
}

case "$SYSENV" in
    dev)
        SSMPATH="/quire/development/msteams/env"
        get_ssm_parameter "$SSMPATH"
        ;;
    prod)
        SSMPATH="/quire/production/msteams/env"
        get_ssm_parameter "$SSMPATH"
        ;;
    *)
        echo "Using bind mount .env file"
        ;;
esac

cd /app && /usr/local/bin/node index.js
