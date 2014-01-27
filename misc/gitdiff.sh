#!/bin/bash
git config --global diff.sheetjs.textconv "j"
touch ~/.gitattributes
cat <<EOF >>~/.gitattributes
*.xlsm diff=sheetjs
*.xlsx diff=sheetjs
*.xls diff=sheetjs
*.XLSM diff=sheetjs
*.XLSX diff=sheetjs
*.XLS diff=sheetjs
EOF
git config --global core.attributesfile '~/.gitattributes'
