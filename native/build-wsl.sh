#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
COMPILER="${CXX:-x86_64-w64-mingw32-g++}"
WINDRES="${WINDRES:-x86_64-w64-mingw32-windres}"
BUILD_TYPE="${1:-Release}"
OUTPUT_DIR="$ROOT_DIR/bin/$BUILD_TYPE"
OUTPUT_EXE="$OUTPUT_DIR/WindowLayouter.Native.exe"
RESOURCE_OBJ="$OUTPUT_DIR/app-icon-res.o"

if ! command -v "$COMPILER" >/dev/null 2>&1; then
  echo "Missing compiler: $COMPILER" >&2
  echo "Install mingw-w64 first, for example:" >&2
  echo "  sudo apt-get update && sudo apt-get install -y mingw-w64" >&2
  exit 1
fi

if ! command -v "$WINDRES" >/dev/null 2>&1; then
  echo "Missing resource compiler: $WINDRES" >&2
  exit 1
fi

mkdir -p "$OUTPUT_DIR"
mkdir -p "$OUTPUT_DIR/lang"

if [[ "${GENERATE_ICONS:-0}" == "1" ]] && command -v python3 >/dev/null 2>&1; then
  python3 "$ROOT_DIR/tools/generate_icon.py" >/dev/null
fi

BUILD_DAY_RAW="$(date +%j)"
BUILD_TIME_RAW="$(date +%H%M)"
BUILD_DAY=$((10#$BUILD_DAY_RAW))
BUILD_TIME=$((10#$BUILD_TIME_RAW))
cat >"$ROOT_DIR/build-version.h" <<EOF
#pragma once

#define WL_VERSION_COMMA 0,3,$BUILD_DAY,$BUILD_TIME
#define WL_VERSION_DOT "0.3.$BUILD_DAY.$BUILD_TIME"
EOF

CXXFLAGS=(
  -std=c++20
  -municode
  -mwindows
  -Wall
  -Wextra
)

LDFLAGS=(
  -static
  -static-libgcc
  -static-libstdc++
  -lcomctl32
  -lshell32
  -lshcore
  -ldwmapi
  -luxtheme
)

if [[ "$BUILD_TYPE" == "Release" ]]; then
  CXXFLAGS+=(-O2 -s)
else
  CXXFLAGS+=(-O0 -g)
fi

echo "Compiler: $COMPILER"
echo "Windres: $WINDRES"
echo "Build type: $BUILD_TYPE"
echo "Output: $OUTPUT_EXE"

(
  cd "$ROOT_DIR"
  "$WINDRES" \
    "WindowLayouter.Native.rc" \
    -O coff \
    -o "$RESOURCE_OBJ"
)

"$COMPILER" \
  "${CXXFLAGS[@]}" \
  "$ROOT_DIR/main.cpp" \
  "$RESOURCE_OBJ" \
  -o "$OUTPUT_EXE" \
  "${LDFLAGS[@]}"

cp -f "$ROOT_DIR"/lang/*.ini "$OUTPUT_DIR/lang/"

ls -lh "$OUTPUT_EXE"
