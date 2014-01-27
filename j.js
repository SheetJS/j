/* j -- (C) 2013-2014 SheetJS -- http://sheetjs.com */
/* vim: set ts=2: */
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
		/* Unknown */
		default: return [undefined, f];
	}
};

function to_json(w) {
	var XL = w[0], workbook = w[1];
	var result = {};
	workbook.SheetNames.forEach(function(sheetName) {
		var roa = XL.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
		if(roa.length > 0) result[sheetName] = roa;
	});
	return result;
}

function to_dsv(w, FS) {
	var XL = w[0], workbook = w[1];
	var result = [];
	workbook.SheetNames.forEach(function(sheetName) {
		var csv = XL.utils.make_csv(workbook.Sheets[sheetName], {FS:FS||","});
		if(csv.length > 0){
			result.push("SHEET: " + sheetName);
			result.push("");
			result.push(csv);
		}
	});
	return result.join("\n");
}

function get_columns(sheet, XL) {
	var val, rowObject, range, columnHeaders, emptyRow, C;
	range = XL.utils.decode_range(sheet["!ref"]);
	columnHeaders = [];
	for (C = range.s.c; C <= range.e.c; ++C) {
		val = sheet[XL.utils.encode_cell({c: C, r: range.s.r})];
		if(val){
			switch(val.t) {
				case 's': case 'str': columnHeaders[C] = val.v; break;
				case 'n': columnHeaders[C] = val.v; break;
			}
		}
	}
	return columnHeaders;
}

function to_html(w) {
	var XL = w[0], wb = w[1];
	var json = to_json(w);
	var tbl = [];
	wb.SheetNames.forEach(function(sheet) {
		var cols = get_columns(wb.Sheets[sheet], XL);
		var src = "<h3>" + sheet + "</h3>";
		src += "<table>";
		src += "<thead><tr>";
		cols.forEach(function(c) { src += "<th>" + (typeof c !== "undefined" ? c : "") + "</th>"; });
		src += "</tr></thead>";
		(json[sheet]||[]).forEach(function(row) {
			src += "<tr>";
			cols.forEach(function(c) { src += "<td>" + (typeof row[c] !== "undefined" ? row[c] : "") + "</c>"; });
			src += "</tr>";
		});
		src += "</table>";
		tbl.push(src);
	});
	return tbl;
};

module.exports = {
	XLSX: XLSX,
	XLS: XLS,
	readFile:readFileSync,
	utils: {
		to_csv: to_dsv,
		to_dsv: to_dsv,
		to_json: to_json,
		to_html: to_html
	},
	version: "XLS " + XLS.version + " ; XLSX " + XLSX.version
};
