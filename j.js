/* j -- (C) 2013-2014 SheetJS -- http://sheetjs.com */
/* vim: set ts=2: */
/*jshint node:true */
var XLSX = require('xlsx');
var XLS = require('xlsjs');
var fs = require('fs');
var readFileSync = function(filename, options) {
	var f = fs.readFileSync(filename);
	switch(f[0]) {
		/* CFB container */
		case 0xd0: return [XLS, XLS.readFile(filename, options)];
		/* XML container (assumed 2003/2004) */
		case 0x3c: return [XLS, XLS.readFile(filename, options)];
		/* Zip container */
		case 0x50: return [XLSX, XLSX.readFile(filename, options)];
		/* Unknown */
		default: return [undefined, f];
	}
};

function to_formulae(w) {
	var XL = w[0], workbook = w[1];
	var result = {};
	workbook.SheetNames.forEach(function(sheetName) {
		var f = XL.utils.get_formulae(workbook.Sheets[sheetName]);
		if(f.length > 0) result[sheetName] = f;
	});
	return result;
}

function to_json(w, raw) {
	var XL = w[0], workbook = w[1];
	var result = {};
	workbook.SheetNames.forEach(function(sheetName) {
		var roa = XL.utils.sheet_to_row_object_array(workbook.Sheets[sheetName], {raw:raw});
		if(roa.length > 0) result[sheetName] = roa;
	});
	return result;
}

function to_dsv(w, FS, RS) {
	var XL = w[0], workbook = w[1];
	var result = {};
	workbook.SheetNames.forEach(function(sheetName) {
		var csv = XL.utils.make_csv(workbook.Sheets[sheetName], {FS:FS||",",RS:RS||"\n"});
		if(csv.length > 0) result[sheetName] = csv;
	});
	return result;
}

function get_cols(sheet, XL) {
	var val, r, hdr, R, C, _XL = XL || XLS;
	hdr = [];
	if(!sheet["!ref"]) return hdr;
	r = _XL.utils.decode_range(sheet["!ref"]);
	for (R = r.s.r, C = r.s.c; C <= r.e.c; ++C) {
		val = sheet[_XL.utils.encode_cell({c:C, r:R})];
		if(!val) continue;
		hdr[C] = typeof val.w !== 'undefined' ? val.w : _XL.utils.format_cell ? XL.utils.format_cell(val) : val.v;
	}
	return hdr;
}

function to_html(w) {
	var XL = w[0], wb = w[1];
	var json = to_json(w);
	var tbl = {};
	wb.SheetNames.forEach(function(sheet) {
		var cols = get_cols(wb.Sheets[sheet], XL);
		var src = "<h3>" + sheet + "</h3>";
		src += "<table>";
		src += "<thead><tr>";
		cols.forEach(function(c) { src += "<th>" + (typeof c !== "undefined" ? c : "") + "</th>"; });
		src += "</tr></thead>";
		(json[sheet]||[]).forEach(function(row) {
			src += "<tr>";
			cols.forEach(function(c) { src += "<td>" + (typeof row[c] !== "undefined" ? row[c] : "") + "</td>"; });
			src += "</tr>";
		});
		src += "</table>";
		tbl[sheet] = src;
	});
	return tbl;
}

var encodings = {
	'&quot;': '"',
	'&apos;': "'",
	'&gt;': '>',
	'&lt;': '<',
	'&amp;': '&'
};
function evert(obj) {
	var o = {};
	Object.keys(obj).forEach(function(k) { if(obj.hasOwnProperty(k)) o[obj[k]] = k; });
	return o;
}
var rencoding = evert(encodings);
var rencstr = "&<>'\"".split("");
function escapexml(text){
	var s = text + '';
	rencstr.forEach(function(y){s=s.replace(new RegExp(y,'g'), rencoding[y]);});
	return s;
}

var cleanregex = /[^A-Za-z0-9_.:]/g;
function to_xml(w) {
	var json = to_json(w);
	var lst = {};
	w[1].SheetNames.forEach(function(sheet) {
		var js = json[sheet], s = sheet.replace(cleanregex,"").replace(/^([0-9])/,"_$1");
		var xml = "";
		xml += "<" + s + ">";
		(js||[]).forEach(function(r) {
			xml += "<" + s + "Data>";
			for(var y in r) if(r.hasOwnProperty(y)) xml += "<" + y.replace(cleanregex,"").replace(/^([0-9])/,"_$1") + ">" + escapexml(r[y]) + "</" +  y.replace(cleanregex,"").replace(/^([0-9])/,"_$1") + ">";
			xml += "</" + s + "Data>";
		});
		xml += "</" + s + ">";
		lst[sheet] = xml;
	});
	return lst;
}

function to_clit(w, s) {
	var t = to_dsv(w, "\u2603","\u2604")[s].split("\u2604").map(function(s) { return s.split("\u2603").map(function(k) { return k.replace(/^"/,"").replace(/"$/,""); }); });
	return t;
}

module.exports = {
	XLSX: XLSX,
	XLS: XLS,
	readFile:readFileSync,
	utils: {
		to_csv: to_dsv,
		to_dsv: to_dsv,
		to_xml: to_xml,
		to_json: to_json,
		to_html: to_html,
		to_formulae: to_formulae,
		to_clit: to_clit,
		get_cols: get_cols
	},
	version: "XLS " + XLS.version + " ; XLSX " + XLSX.version
};
