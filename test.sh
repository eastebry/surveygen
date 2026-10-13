#!/bin/sh
set -eu
ROOT=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
# Build the PDFs first: the suite also checks the published sample's layout.
"$ROOT/build.sh"
docker run --rm --network none --user "$(id -u):$(id -g)" -e HOME=/tmp \
    -v "$ROOT:/work" --entrypoint python3 community-survey:local \
    -m unittest discover -s tests -v
