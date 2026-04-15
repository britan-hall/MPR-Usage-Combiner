#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT_DIR"

mvn -q -DskipTests package

java -jar "target/mpr-combiner-1.0.0-all.jar" \
  --input "data/input" \
  --recursive true \
  --output "data/output/combined-usage.xlsx" \
  "$@"

echo "Wrote: data/output/combined-usage.xlsx"

