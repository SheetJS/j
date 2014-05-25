#!/usr/bin/env osascript
-- open_excel_2011.scpt -- open file using Excel 2011 for Mac
-- Copyright (C) 2014  SheetJS
-- vim: set ts=2:

on run argv
	set pwd to (system attribute "PWD")
	set workingDir to POSIX path of pwd
	set input_file_name to pwd & "/" & (item 1 of argv)
	set input_file to POSIX file input_file_name
	tell application "Microsoft Excel"
		open workbook workbook file name input_file update links do not update links read only true ignore read only recommended true notify false add to mru false
		try
			tell active workbook to close
		end try
		quit saving no
	end tell
end run
