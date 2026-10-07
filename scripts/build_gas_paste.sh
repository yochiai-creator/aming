#!/bin/sh
# gas/ を Apps Script エディタに貼り付ける用の3ファイルにまとめる → gas-paste/
cd "$(dirname "$0")/.." && mkdir -p gas-paste
for f in gas/01_*.js gas/02_*.js gas/03_*.js gas/04_*.js; do echo "/* ===== ${f#gas/} ===== */"; grep -v "^if (typeof module" "$f"; echo; done > gas-paste/コード.gs
cp gas/index.html gas/appsscript.json gas-paste/
