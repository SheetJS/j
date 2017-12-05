/* j -- (C) 2013-present  SheetJS -- http://sheetjs.com */
/* vim: set ts=2: */
/*jshint node:true, eqnull:true */
var X = require('xlsx');

var readFileSync = function(filename/*:string*/, options/*:any*/)/*:JWorkbook*/ { return [X, X.readFile(filename, options)]; };
var read = function(data/*:any*/, options/*:any*/)/*:JWorkbook*/ { return [X, X.read(data, options)]; };

function to_formulae(w/*:JWorkbook*/) {
	var workbook = w[1];
	var result = {};
	workbook.SheetNames.forEach(function(sheetName) {
		var f = X.utils.get_formulae(workbook.Sheets[sheetName]);
		if(f.length > 0) result[sheetName] = f;
	});
	return result;
}

function to_json(w/*:JWorkbook*/, raw/*:?boolean*/) {
	var workbook = w[1];
	var result = {};
	workbook.SheetNames.forEach(function(sheetName) {
		var roa = X.utils.sheet_to_json(workbook.Sheets[sheetName], typeof raw == "object" ? raw : {raw:raw});
		if(roa.length > 0) result[sheetName] = roa;
	});
	return result;
}

function to_dsv(w/*:JWorkbook*/, FS/*:?string*/, RS/*:?string*/) {
	var workbook = w[1];
	var result = {};
	workbook.SheetNames.forEach(function(sheetName) {
		var csv = X.utils.make_csv(workbook.Sheets[sheetName], {FS:FS||",",RS:RS||"\n"});
		if(csv.length > 0) result[sheetName] = csv;
	});
	return result;
}

function get_cols(sheet/*:Worksheet*/) {
	var val, r, hdr, R, C;
	hdr = [];
	if(!sheet["!ref"]) return hdr;
	r = X.utils.decode_range(sheet["!ref"]);
	for (R = r.s.r, C = r.s.c; C <= r.e.c; ++C) {
		val = sheet[X.utils.encode_cell({c:C, r:R})];
		if(val == null) continue;
		hdr[C] = val.w !== undefined ? val.w : X.utils.format_cell ? X.utils.format_cell(val) : val.v;
	}
	return hdr;
}

function to_md(w/*:JWorkbook*/) {
	var wb = w[1];
	var tbl = {};
	wb.SheetNames.forEach(function(sheet) {
		var ws = wb.Sheets[sheet];
		if(ws["!ref"] == null) return;
		var src = "|", val, w;
		var range = X.utils.decode_range(ws["!ref"]);
		var R = range.s.r, C;
		for(C = range.s.c; C <= range.e.c; ++C) {
			val = ws[X.utils.encode_cell({c:C,r:R})];
			w = val == null ? "" : val.w !== undefined ? val.w : X.utils.format_cell ? X.utils.format_cell(val) : val.v;
			src += w + "|";
		}
		src += "\n|";
		for(C = range.s.c; C <= range.e.c; ++C) {
			val = ws[X.utils.encode_cell({c:C,r:R})];
			w = val == null ? "" : val.w !== undefined ? val.w : X.utils.format_cell ? X.utils.format_cell(val) : val.v;
			src += " ---- |";
		}
		src += "\n";
		for(R = range.s.r+1; R <= range.e.r; ++R) {
			src += "|";
			for(C = range.s.c; C <= range.e.c; ++C) {
				val = ws[X.utils.encode_cell({c:C,r:R})];
				w = val == null ? "" : val.w !== undefined ? val.w : X.utils.format_cell ? X.utils.format_cell(val) : val.v;
				src += w + "|";
			}
			src += "\n";
		}
		tbl[sheet] = src;
	});
	return tbl;
}

function to_html(w/*:JWorkbook*/) {
	var wb = w[1];
	var tbl = {};
	wb.SheetNames.forEach(function(sheet) {
		var ws = wb.Sheets[sheet];
		if(ws["!ref"] == null) return;
		var src = "<h3>" + sheet + "</h3>";
		var range = X.utils.decode_range(ws["!ref"]);
		src += "<table>";
		src += "<colgroup span=\"" + (range.e.c - range.s.c + 1) + "\"></colgroup>";
		for(var R = range.s.r; R <= range.e.r; ++R) {
			src += "<tr>";
			for(var C = range.s.c; C <= range.e.c; ++C) {
				var val = ws[X.utils.encode_cell({c:C,r:R})];
				var w = val == null ? "" : val.w !== undefined ? val.w : X.utils.format_cell ? X.utils.format_cell(val) : val.v;
				src += "<td>" + w + "</td>";
			}
			src += "</tr>";
		}
		src += "</table>";
		tbl[sheet] = src;
	});
	return tbl;
}

function to_html_cols(w/*:JWorkbook*/) {
	var wb = w[1];
	var json = to_json(w);
	var tbl = {};
	wb.SheetNames.forEach(function(sheet) {
		var cols = get_cols(wb.Sheets[sheet]);
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
function evert(obj/*:Object*/)/*:Object*/ {
	var o = {};
	Object.keys(obj).forEach(function(k) { if(obj.hasOwnProperty(k)) o[obj[k]] = k; });
	return o;
}
var rencoding = evert(encodings);
var rencstr = "&<>'\"".split("");
function escapexml(text/*:string*/)/*:string*/{
	var s = text + '';
	rencstr.forEach(function(y){s=s.replace(new RegExp(y,'g'), rencoding[y]);});
	return s;
}

var cleanregex = /[^A-Za-z0-9_.:]/g;
function to_xml(w/*:JWorkbook*/) {
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

function to_wb(t/*:string*/) {
	return function(w/*:JWorkbook*/, o/*:any*/) {
		o = o || {}; o.bookType = t;
		if(o.bookSST === undefined) o.bookSST = true;
		if(o.type === undefined) o.type = 'buffer';
		return X.write(w[1], o);
	};
}

function to_ws(t/*:string*/) {
	return function(w/*:JWorkbook*/, o/*:any*/) {
		var workbook = w[1];
		var result = {};
		o = o || {}; o.bookType = t;
		if(o.bookSST === undefined) o.bookSST = true;
		if(o.type === undefined) o.type = 'buffer';
		workbook.SheetNames.forEach(function(sheetName) {
			var ws = workbook.Sheets[sheetName];
			if(!ws || !ws['!ref']) return;
			o.sheet = sheetName;
			try {
				var out = X.write(workbook, o);
				result[sheetName] = out;
			} catch(e) { console.error(e); }
		});
		return result;
	};
}

function util(t) {
	return function(w/*:JWorkbook*/) {
		var workbook = w[1];
		var result = {};
		workbook.SheetNames.forEach(function(sheetName) {
			var ws = workbook.Sheets[sheetName];
			if(!ws || !ws['!ref']) return;
			try {
				var out = X.utils["sheet_to_" + t](ws);
				result[sheetName] = out;
			} catch(e) { console.error(e); }
		});
		return result;
	};
}

var utils = {
	to_csv: to_dsv,
	to_dsv: to_dsv,
	to_xml: to_xml,
	to_xlsx: to_wb('xlsx'),
	to_xlsm: to_wb('xlsm'),
	to_xlsb: to_wb('xlsb'),
	to_ods: to_wb('ods'),
	to_fods: to_wb('fods'),
	to_biff2: to_ws('biff2'),
	to_biff5: to_wb('biff5'),
	to_biff8: to_wb('biff8'),
	to_xlml: to_wb('xlml'),
	to_xls: to_wb('biff8'),
	to_dbf: to_ws('dbf'),
	to_txt: to_ws('txt'),
	to_rtf: to_ws('rtf'),
	to_prn: to_ws('prn'),
	to_json: to_json,
	to_html: to_html,
	to_html_cols: to_html_cols,
	to_formulae: to_formulae,
	to_md: to_md,
	to_dif: util('dif'),
	to_sylk: util('slk'),
	to_eth: util('eth'),
	to_socialcalc: util('eth'),
	get_cols: get_cols
};
var J = ({
	XLSX: X,
	XLS: X,
	readFile:readFileSync,
	read:read,
	utils: utils,
	version: "XLSX " + X.version
}/*:any*/);

if(typeof module !== 'undefined') module.exports = J;
