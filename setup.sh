#!/bin/bash
# setup.sh — Run this at the start of each Cowork session.
# Finds the mounted Cowork folder, copies all pipeline files to
# /tmp/pipeline_data/, verifies the script, and prints the path for step 7.
#
# Usage: bash setup.sh

set -e

DEST="/tmp/pipeline_data"

# Find the mounted Cowork folder by looking for daily_pipeline.py
# (works regardless of what the mount path is called in this sandbox)
COWORK_DIR=""
for candidate in \
    /mnt/user/Cowork \
    /mnt/Cowork \
    /sessions/*/mnt/Cowork \
    /home/user/Cowork \
    ~/Cowork \
    /workspace \
    /mnt/workspace \
    /mnt/user-data/uploads \
; do
    # Expand glob
    for path in $candidate; do
        if [ -f "$path/daily_pipeline.py" ]; then
            COWORK_DIR="$path"
            break 2
        fi
    done
done

if [ -z "$COWORK_DIR" ]; then
    echo "ERROR: Could not find daily_pipeline.py in any expected mount location."
    echo "Searched: /mnt/user/Cowork, /mnt/Cowork, /sessions/*/mnt/Cowork, etc."
    echo "If your mount path is different, run manually:"
    echo "  cp -r /your/actual/path/* /tmp/pipeline_data/"
    exit 1
fi

echo "Found Cowork folder: $COWORK_DIR"
echo "Copying to $DEST ..."
mkdir -p "$DEST"
cp -r "$COWORK_DIR"/. "$DEST/"

cd "$DEST"

echo ""
echo "Verifying script..."
LINES=$(wc -l < daily_pipeline.py)
python3 -m py_compile daily_pipeline.py && echo "  Script OK: $LINES lines"

if [ "$LINES" -lt 1700 ]; then
    echo ""
    echo "WARNING: daily_pipeline.py has only $LINES lines (expected 1700+)."
    echo "The file may be truncated. Stop and re-copy from the Claude.ai conversation."
    exit 1
fi

if [ ! -f greenhouse_api.py ]; then
    echo ""
    echo "WARNING: greenhouse_api.py is missing. daily_pipeline.py imports it"
    echo "and will fail immediately. Re-copy it from the Claude.ai conversation."
    exit 1
fi
python3 -c "import sys; sys.path.insert(0, '.'); import daily_pipeline" \
    && echo "  Import OK: daily_pipeline.py + greenhouse_api.py both load cleanly" \
    || { echo "WARNING: daily_pipeline.py failed to import - check the error above."; exit 1; }

echo ""
echo "Setup complete. Run all pipeline commands from: $DEST"
echo ""
echo "When done, copy results back with:"
echo "  cp $DEST/*.xlsx $DEST/*.json \"$COWORK_DIR/\" 2>/dev/null && echo 'Copy done'"
