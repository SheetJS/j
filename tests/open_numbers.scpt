#!/usr/bin/env osascript
-- open_numbers.scpt -- open file using Numbers for Mac
-- Copyright (C) 2014  SheetJS
-- vim: set ts=2:

on run argv
	set pwd to (system attribute "PWD")
	set workingDir to POSIX path of pwd
	set input_file_name to pwd & "/" & (item 1 of argv)
	set input_file to POSIX file input_file_name
	tell application "Numbers"
		open input_file 
		try
			close saving no
		end try
		quit saving no
	end tell
end run
