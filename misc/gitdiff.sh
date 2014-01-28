#!/bin/bash

# Define a sheetjs diff type that uses j
git config --global diff.sheetjs.textconv "j"

# Configure a user .gitattributes file that maps the xls{,x,m} files
touch ~/.gitattributes
cat <<EOF >>~/.gitattributes
*.xls diff=sheetjs
*.xlsb diff=sheetjs
*.xlsm diff=sheetjs
*.xlsx diff=sheetjs
*.XLS diff=sheetjs
*.XLSB diff=sheetjs
*.XLSM diff=sheetjs
*.XLSX diff=sheetjs
EOF

# Set the .gitattributes to be used for all repos on the system:
git config --global core.attributesfile '~/.gitattributes'
