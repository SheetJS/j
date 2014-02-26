#!/bin/bash
if [ ! -e test_files ]; then git clone https://github.com/SheetJS/test_files; fi
cd test_files
git pull
make init
