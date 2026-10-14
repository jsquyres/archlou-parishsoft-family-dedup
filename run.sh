#!/usr/bin/env bash
# Run dedup.py on the ParishSoft exports in $DATA_DIR and write reports and
# caches to $DATA_DIR/output.
#
# DATA_DIR must be OUTSIDE this git repository: the exports and reports
# contain personal data.  Default: ~/archlou-dedup.
#
# Override inputs with environment variables, e.g.:
#   DATA_DIR=/secure/dir ./run.sh
#   FAMILIES=/path/families.xlsx MEMBERS=/path/members.xlsx ./run.sh
# Extra arguments are passed through to dedup.py.
set -euo pipefail
repo=$(cd "$(dirname "$0")" && pwd)

PYTHON=${PYTHON:-python3}
DATA_DIR=${DATA_DIR:-$HOME/archlou-dedup}
FAMILIES=${FAMILIES:-$DATA_DIR/archlou-families-list-to-deduplicate.xlsx}
MEMBERS=${MEMBERS:-$DATA_DIR/archlou-members-list.xlsx}
ENHANCED=${ENHANCED:-$DATA_DIR/archlou-families-list-to-deduplicate-enhanced.xlsx}
OUT_DIR=${OUT_DIR:-$DATA_DIR/output}

case "$(cd "$DATA_DIR" 2>/dev/null && pwd)/" in
    "$repo"/*) echo "error: DATA_DIR ($DATA_DIR) is inside the git repo; use a directory outside it" >&2; exit 1 ;;
esac

args=("$FAMILIES" "$MEMBERS" --out-dir "$OUT_DIR")
if [[ -f "$ENHANCED" ]]; then
    args+=(--enhanced "$ENHANCED")
else
    echo "note: $ENHANCED not found; running without enhanced Family fields" >&2
fi

exec "$PYTHON" -I "$repo/dedup.py" "${args[@]}" "$@"
