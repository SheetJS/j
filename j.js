/* j -- (C) 2013-2014 SheetJS -- http://sheetjs.com */
/* vim: set ts=2: */
/*jshint node:true, eqnull:true */
var XLSX = require('xl'+'sx');
var XLS = require('xl'+'sjs');
var HARB = require('ha'+'rb');
var UTILS = XLSX.utils;

var libs = [
	["XLS", XLS],
	["XLSX", XLSX],
	["HARB", HARB]
];

var fs = require('f'+'s');
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
		default: return [HARB, HARB.readFile(filename, options)];
	}
};

var read = function(data, options) {
	switch(data[0]) {
		/* CFB container */
		case 0xd0: return [XLS, XLS.read(data, options)];
		/* XML container (assumed 2003/2004) */
		case 0x3c: return [XLS, XLS.read(data, options)];
		/* Zip container */
		case 0x50: return [XLSX, XLSX.read(data, options)];
		/* Unknown */
		default: return [HARB, HARB.read(data.toString(), options)];
	}
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
		var roa = XL.utils.sheet_to_row_object_array(workbook.Sheets[sheetName], {raw:raw});
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

/* originally from http://git.io/xlsx2socialcalc */
/* xlsx2socialcalc.js (C) 2014 SheetJS -- http://sheetjs.com */
var sheet_to_socialcalc = (function() {
	var header = [
		"socialcalc:version:1.5",
		"MIME-Version: 1.0",
		"Content-Type: multipart/mixed; boundary=SocialCalcSpreadsheetControlSave"
	].join("\n");

	var sep = [
		"--SocialCalcSpreadsheetControlSave",
		"Content-type: text/plain; charset=UTF-8",
		""
	].join("\n");

	/* TODO: the other parts */
	var meta = [
		"# SocialCalc Spreadsheet Control Save",
		"part:sheet"
	].join("\n");

	var end = "--SocialCalcSpreadsheetControlSave--";

	var scencode = function(s) { return s.replace(/\\/g, "\\b").replace(/:/g, "\\c").replace(/\n/g,"\\n"); };

	var scsave = function scsave(ws) {
		if(!ws || !ws['!ref']) return "";
		var o = [], oo = [], cell, coord;
		var r = UTILS.decode_range(ws['!ref']);
		for(var R = r.s.r; R <= r.e.r; ++R) {
			for(var C = r.s.c; C <= r.e.c; ++C) {
				coord = UTILS.encode_cell({r:R,c:C});
				if(!(cell = ws[coord]) || cell.v == null) continue;
				oo = ["cell", coord, 't'];
				switch(cell.t) {
					case 's': case 'str': oo.push(scencode(cell.v)); break;
					case 'n':
						if(cell.f) {
							oo[2] = 'vtf';
							oo.push('n');
							oo.push(cell.v);
							oo.push(scencode(cell.f));
						}
						else {
							oo[2] = 'v';
							oo.push(cell.v);
						} break;
					case 'b':
						if(cell.f) {
							oo[2] = 'vtf';
							oo.push('nl');
							oo.push(cell.v ? 1 : 0);
							oo.push(scencode(cell.f));
						} else {
							oo[2] = 'vtc';
							oo.push('nl');
							oo.push(cell.v ? 1 : 0);
							oo.push(cell.v ? 'TRUE' : 'FALSE');
						} break;
				}
				o.push(oo.join(":"));
			}
		}
		o.push("sheet:c:" + (r.e.c - r.s.c + 1) + ":r:" + (r.e.r - r.s.r + 1) + ":tvf:1");
		o.push("valueformat:1:text-wiki");
		o.push("copiedfrom:" + ws['!ref']);
		return o.join("\n");
	};

	return function socialcalcify(ws, opts) {
		return [header, sep, meta, sep, scsave(ws), end].join("\n");
		// return ["version:1.5", scsave(ws)].join("\n"); // clipboard form
	};
})();

function to_socialcalc(w) {
	var workbook = w[1];
	var result = {};
	workbook.SheetNames.forEach(function(sheetName) {
		var socialcalc = sheet_to_socialcalc(workbook.Sheets[sheetName]);
		if(socialcalc.length > 0) result[sheetName] = socialcalc;
	});
	return result;
}

var version = libs.map(function(x) { return x[0] + " " + x[1].version; }).join(" ; ");

var utils = {
	to_csv: to_dsv,
	to_dsv: to_dsv,
	to_xml: to_xml,
	to_xlsx: to_xlsx,
	to_xlsm: to_xlsm,
	to_xlsb: to_xlsb,
	to_json: to_json,
	to_html: to_html,
	to_html_cols: to_html_cols,
	to_formulae: to_formulae,
	to_md: to_md,
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
if(typeof process !== 'undefined') {
	/* see https://github.com/SheetJS/j/issues/4 */
	if(process.version === 'v0.10.31') {
		var msgs = [
			"node v0.10.31 is known to crash on OSX and Linux, refusing to proceed.",
			"see https://github.com/SheetJS/j/issues/4 for the relevant discussion.",
			"see https://github.com/joyent/node/issues/8208 for the relevant node issue"
		];
		msgs.forEach(function(m) { console.error(m); });
		throw "node v0.10.31 bug";
	}
}
