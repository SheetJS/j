/* Copyright (C) 2013  SheetJS */
var XLSX = require('xlsx');
var XLS = require('xlsjs');
var fs = require('fs');
var readFileSync = function(filename, options) {
	var f = fs.readFileSync(filename);
	switch(f[0]) {
		/* CFB container */
		case 0xd0: return [XLS, XLS.readFile(filename)];
		/* Zip container */
		case 0x50: return [XLSX, XLSX.readFile(filename)];
	}
	return [undefined, f];
};
module.exports = { XLSX: XLSX, XLS: XLS, readFile:readFileSync };
