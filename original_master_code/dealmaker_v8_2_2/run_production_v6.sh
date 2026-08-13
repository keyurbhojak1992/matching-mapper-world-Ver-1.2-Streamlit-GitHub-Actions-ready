#!/usr/bin/env sh
set -eu
if [ "$#" -lt 3 ]; then
  echo "Usage: $0 SALES ACTIONS STOCK [CONTROL] [OUTPUT_DIR]" >&2
  exit 2
fi
SCRIPT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
CONTROL=${4:-$SCRIPT_DIR/Dealmaker_Full_Control_v6.xlsx}
OUTPUT=${5:-$SCRIPT_DIR/dealmaker_v8_2_2_output}
exec python3 "$SCRIPT_DIR/Dealmaker_Production_Runner_v6.py" \
  --sales "$1" --actions "$2" --stock "$3" --control "$CONTROL" \
  --sales-header-row 0 --actions-header-row 0 --stock-header-row 0 \
  --output-dir "$OUTPUT"
