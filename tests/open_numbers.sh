#!/bin/bash
# open_numbers.sh -- open every generated test file in Numbers for Mac 
# Copyright (C) 2014  SheetJS
# vim: set ts=2:
timeout() { perl -e 'alarm shift; exec @ARGV; ' "$@"; }
cnt=0
for i in test_files/*__.xlsx; do
	echo "$i"
	timeout 5 osascript -s o ./tests/open_numbers.scpt "$i"
	killall -9 Numbers &>/dev/null 
	sleep 1
	if ((++cnt > 10)); then
		((cnt=0))
		sleep 4
	fi
done
