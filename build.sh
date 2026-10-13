#!/bin/sh
set -eu
ROOT=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
IMAGE=community-survey:local
if ! command -v docker >/dev/null 2>&1; then
    echo 'Docker is required. Install Docker and start its engine, then rerun ./build.sh.' >&2
    exit 1
fi
if ! docker info >/dev/null 2>&1; then
    echo 'Cannot access the Docker engine. Start Docker and check your Docker permissions.' >&2
    exit 1
fi
mkdir -p "$ROOT/output/pdf" "$ROOT/build"
docker build -t "$IMAGE" "$ROOT"
docker run --rm --network none --user "$(id -u):$(id -g)" \
    -e HOME=/tmp -e SURVEY_LANGUAGE="${SURVEY_LANGUAGE:-all}" \
    -v "$ROOT:/work" "$IMAGE" "$@"
