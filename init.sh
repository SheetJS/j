#!/bin/bash
mkdir -p tmp
mkdir -p tmp/2011 tmp/2013 tmp/2016
mkdir -p tmp/biff{2,3,4,5}
mkdir -p tmp/artifacts/{wps,quattro}
if [ ! -e test_files ]; then git clone https://github.com/SheetJS/test_files; fi
cd test_files
git pull
make init
