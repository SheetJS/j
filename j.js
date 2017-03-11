/* j -- (C) 2013-present  SheetJS -- http://sheetjs.com */
/* vim: set ts=2: */
/*jshint node:true, eqnull:true */
var XLSX = require('xlsx');
var XLS = XLSX;
var HARB = require('harb');
var UTILS = XLSX.utils;

var libs = [
	["XLS", XLS],
	["XLSX", XLSX],
	["HARB", HARB]
];

var fs = require('fs');

var is_xlsx = function(d) {
	switch(d[0]) {
		/* CFB container */
		case 0xd0: return true;
		/* XML container (assumed 2003/2004) */
		case 0x3c: return true;
		/* BIFF */
		case 0x09: return true;
		/* Zip container or plaintext */
		case 0x50: return (d[1] == 0x4b && d[2] <= 0x10 && d[3] <= 0x10);
		/* Unknown */
		default: return false;
	}
}
var readFileSync = function(filename, options) {
	var f = fs.readFileSync(filename);
	if(is_xlsx(f)) return [XLSX, XLSX.readFile(filename, options)];
	else return [HARB, HARB.readFile(filename, options)];
};

var read = function(data, options) {
	if(is_xlsx(data)) return [XLSX, XLSX.read(data, options)];
	else return [HARB, HARB.read(data.toString(), options)];
};

function to_formulae(w) {
	var XL = w[0], workbook = w[1];
	if(!XL.utils.get_formulae) XL = XLSX;
	var result = {};
	workbook.SheetNames.forEach(function(sheetName) {
		var f = XL.utils.get_formulae(workbook.Sheets[sheetName]);
		if(f.length > 0) result[sheetName] = f;
	});
	return result;
}

function to_json(w, raw) {
	var XL = w[0], workbook = w[1];
	if(!XL.utils.sheet_to_row_object_array) XL = XLSX;
	var result = {};
	workbook.SheetNames.forEach(function(sheetName) {
		var roa = XL.utils.sheet_to_row_object_array(workbook.Sheets[sheetName], typeof raw == "object" ? raw : {raw:raw});
		if(roa.length > 0) result[sheetName] = roa;
	});
	return result;
}

function to_dsv(w, FS, RS) {
	var XL = w[0], workbook = w[1];
	if(!XL.utils.make_csv) XL = XLSX;
	var result = {};
	workbook.SheetNames.forEach(function(sheetName) {
		var csv = XL.utils.make_csv(workbook.Sheets[sheetName], {FS:FS||",",RS:RS||"\n"});
		if(csv.length > 0) result[sheetName] = csv;
	});
	return result;
}

function get_cols(sheet, XL) {
	var val, r, hdr, R, C, _XL = XL || XLS;
	if(!_XL.utils.format_cell) _XL = XLSX;
	hdr = [];
	if(!sheet["!ref"]) return hdr;
	r = _XL.utils.decode_range(sheet["!ref"]);
	for (R = r.s.r, C = r.s.c; C <= r.e.c; ++C) {
		val = sheet[_XL.utils.encode_cell({c:C, r:R})];
		if(val == null) continue;
		hdr[C] = val.w !== undefined ? val.w : _XL.utils.format_cell ? _XL.utils.format_cell(val) : val.v;
	}
	return hdr;
}

function to_md(w) {
	var XL = w[0], wb = w[1];
	if(!XL.utils.format_cell) XL = XLSX;
	var tbl = {};
	wb.SheetNames.forEach(function(sheet) {
		var ws = wb.Sheets[sheet];
		if(ws["!ref"] == null) return;
		var src = "|", val, w;
		var range = XL.utils.decode_range(ws["!ref"]);
		var R = range.s.r, C;
		for(C = range.s.c; C <= range.e.c; ++C) {
			val = ws[XL.utils.encode_cell({c:C,r:R})];
			w = val == null ? "" : val.w !== undefined ? val.w : XL.utils.format_cell ? XL.utils.format_cell(val) : val.v;
			src += w + "|";
		}
		src += "\n|";
		for(C = range.s.c; C <= range.e.c; ++C) {
			val = ws[XL.utils.encode_cell({c:C,r:R})];
			w = val == null ? "" : val.w !== undefined ? val.w : XL.utils.format_cell ? XL.utils.format_cell(val) : val.v;
			src += " ---- |";
		}
		src += "\n";
		for(R = range.s.r+1; R <= range.e.r; ++R) {
			src += "|";
			for(C = range.s.c; C <= range.e.c; ++C) {
				val = ws[XL.utils.encode_cell({c:C,r:R})];
				w = val == null ? "" : val.w !== undefined ? val.w : XL.utils.format_cell ? XL.utils.format_cell(val) : val.v;
				src += w + "|";
			}
			src += "\n";
		}
		tbl[sheet] = src;
	});
	return tbl;
}

function to_html(w) {
	var XL = w[0], wb = w[1];
	if(!XL.utils.format_cell) XL = XLSX;
	var tbl = {};
	wb.SheetNames.forEach(function(sheet) {
		var ws = wb.Sheets[sheet];
		if(ws["!ref"] == null) return;
		var src = "<h3>" + sheet + "</h3>";
		var range = XL.utils.decode_range(ws["!ref"]);
		src += "<table>";
		src += "<colgroup span=\"" + (range.e.c - range.s.c + 1) + "\"></colgroup>";
		for(var R = range.s.r; R <= range.e.r; ++R) {
			src += "<tr>";
			for(var C = range.s.c; C <= range.e.c; ++C) {
				var val = ws[XL.utils.encode_cell({c:C,r:R})];
				var w = val == null ? "" : val.w !== undefined ? val.w : XL.utils.format_cell ? XL.utils.format_cell(val) : val.v;
				src += "<td>" + w + "</td>";
			}
			src += "</tr>";
		}
		src += "</table>";
		tbl[sheet] = src;
	});
	return tbl;
}

function to_html_cols(w) {
	var XL = w[0], wb = w[1];
	var json = to_json(w);
	var tbl = {};
	wb.SheetNames.forEach(function(sheet) {
		var cols = get_cols(wb.Sheets[sheet], XL);
		var src = "<h3>" + sheet + "</h3>";
		src += "<table>";
		src += "<thead><tr>";
		cols.forEach(function(c) { src += "<th>" + (c !== undefined ? c : "") + "</th>"; });
		src += "</tr></thead>";
		(json[sheet]||[]).forEach(function(row) {
			src += "<tr>";
			cols.forEach(function(c) { src += "<td>" + (row[c] !== undefined ? row[c] : "") + "</td>"; });
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

function to_xlsx_factory(t) {
	return function(w, o) {
		o = o || {}; o.bookType = t;
		if(o.bookSST === undefined) o.bookSST = true;
		if(o.type === undefined) o.type = 'buffer';
		return XLSX.write(w[1], o);
	};
}

var to_xlsx = to_xlsx_factory('xlsx');
var to_xlsm = to_xlsx_factory('xlsm');
var to_xlsb = to_xlsx_factory('xlsb');
var to_ods = to_xlsx_factory('ods');
var to_fods = to_xlsx_factory('fods');
var to_biff2 = to_xlsx_factory('biff2');

function to_harb_factory(t) {
	return function (w) {
		var workbook = w[1];
		var result = {};
		workbook.SheetNames.forEach(function(sheetName) {
			var out = HARB.utils["sheet_to_" + t](workbook.Sheets[sheetName]);
			result[sheetName] = out;
		});
		return result;
	};
}

var to_dif = to_harb_factory('dif');
var to_sylk = to_harb_factory('sylk');
var to_socialcalc = to_harb_factory('socialcalc');


var version = libs.map(function(x) { return x[0] + " " + x[1].version; }).join(" ; ");

var utils = {
	to_csv: to_dsv,
	to_dsv: to_dsv,
	to_xml: to_xml,
	to_xlsx: to_xlsx,
	to_xlsm: to_xlsm,
	to_xlsb: to_xlsb,
	to_ods: to_ods,
	to_fods: to_fods,
	to_biff2: to_biff2,
	to_json: to_json,
	to_html: to_html,
	to_html_cols: to_html_cols,
	to_formulae: to_formulae,
	to_md: to_md,
	to_dif: to_dif,
	to_sylk: to_sylk,
	to_socialcalc: to_socialcalc,
	get_cols: get_cols
};
var J = {
	XLSX: XLSX,
	XLS: XLS,
	readFile:readFileSync,
	read:read,
	utils: utils,
	version: version
};

if(typeof module !== 'undefined') module.exports = J;
