#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT_DIR"

mvn -q -DskipTests package

DEFAULT_INPUT="data/input"
DEFAULT_OUTPUT="data/output/combined-usage.xlsx"

has_arg() {
  local needle="$1"
  shift
  for a in "$@"; do
    if [[ "$a" == "$needle" ]]; then
      return 0
    fi
  done
  return 1
}

ARGS=()

if ! has_arg "--input" "$@"; then
  ARGS+=(--input "$DEFAULT_INPUT")
fi

if ! has_arg "--output" "$@"; then
  ARGS+=(--output "$DEFAULT_OUTPUT")
fi

ARGS+=(--recursive true)

java -jar "target/mpr-combiner-1.0.0-all.jar" "${ARGS[@]}" "$@"

if has_arg "--output" "$@"; then
  # Best-effort: print the value after --output
  out=""
  prev=""
  for a in "$@"; do
    if [[ "$prev" == "--output" ]]; then
      out="$a"
      break
    fi
    prev="$a"
  done
  if [[ -n "$out" ]]; then
    echo "Wrote: $out"
  fi
else
  echo "Wrote: $DEFAULT_OUTPUT"
fi

