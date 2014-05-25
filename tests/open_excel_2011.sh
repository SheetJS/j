#!/bin/bash
# open_excel_2011.sh -- open every generated test file in Excel 2011 
# Copyright (C) 2014  SheetJS
# vim: set ts=2:
timeout() { perl -e 'alarm shift; exec @ARGV; ' "$@"; }
cnt=0
for i in test_files/*__.xlsx; do
	echo "$i"
	timeout 5 osascript -s o ./tests/open_excel_2011.scpt "$i"
	killall -9 Microsoft\ Excel
	sleep 1
	if ((++cnt > 10)); then
		((cnt=0))
		sleep 4
	fi
done
