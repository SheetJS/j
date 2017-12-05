#!/usr/bin/env node
/* j (C) 2013-present  SheetJS -- http://sheetjs.com */
/* eslint-env node */
var n = "j";
/* vim: set ts=2 ft=javascript: */
var X/*:JModule*/ = (require('../')/*:any*/);
require('exit-on-epipe');
var fs = require('fs'), program = require('commander');
program
	.version(X.version)
	.usage('[options] <file> [sheetname]')
	.option('-f, --file <file>', 'use specified workbook (- for stdin)')
	.option('-s, --sheet <sheet>', 'print specified sheet (default first sheet)')
	.option('-N, --sheet-index <idx>', 'use specified sheet index (0-based)')
	.option('-p, --password <pw>', 'if file is encrypted, try with specified pw')
	.option('-l, --list-sheets', 'list sheet names and exit')
	.option('-o, --output <file>', 'output to specified file')

	.option('-B, --xlsb', 'emit XLSB to <sheetname> or <file>.xlsb')
	.option('-M, --xlsm', 'emit XLSM to <sheetname> or <file>.xlsm')
	.option('-X, --xlsx', 'emit XLSX to <sheetname> or <file>.xlsx')
	.option('-Y, --ods',  'emit ODS  to <sheetname> or <file>.ods')
	.option('-8, --xls',  'emit XLS  to <sheetname> or <file>.xls (BIFF8)')
	.option('-5, --biff5','emit XLS  to <sheetname> or <file>.xls (BIFF5)')
	//.option('-4, --biff4','emit XLS  to <sheetname> or <file>.xls (BIFF4)')
	//.option('-3, --biff3','emit XLS  to <sheetname> or <file>.xls (BIFF3)')
	.option('-2, --biff2','emit XLS  to <sheetname> or <file>.xls (BIFF2)')
	.option('-6, --xlml', 'emit SSML to <sheetname> or <file>.xls (2003 XML)')
	.option('-T, --fods', 'emit FODS to <sheetname> or <file>.fods (Flat ODS)')

	.option('-S, --formulae', 'emit list of values and formulae')
	.option('-j, --json',     'emit formatted JSON (all fields text)')
	.option('-J, --raw-js',   'emit raw JS object (raw numbers)')
	.option('-A, --arrays',   'emit rows as JS objects (raw numbers)')
	.option('-H, --html', 'emit HTML to <sheetname> or <file>.html')
	.option('-D, --dif',  'emit DIF  to <sheetname> or <file>.dif (Lotus DIF)')
	.option('-U, --dbf',  'emit DBF  to <sheetname> or <file>.dbf (MSVFP DBF)')
	.option('-K, --sylk', 'emit SYLK to <sheetname> or <file>.slk (Excel SYLK)')
	.option('-P, --prn',  'emit PRN  to <sheetname> or <file>.prn (Lotus PRN)')
	.option('-E, --socialcalc',  'emit ETH  to <sheetname> or <file>.eth (Ethercalc)')
	.option('-t, --txt',  'emit TXT  to <sheetname> or <file>.txt (UTF-8 TSV)')
	.option('-r, --rtf',  'emit RTF  to <sheetname> or <file>.txt (Table RTF)')
	.option('-x, --xml',  'emit XML')
	.option('-m, --markdown', 'emit markdown table')

	.option('-F, --field-sep <sep>', 'CSV field separator', ",")
	.option('-R, --row-sep <sep>', 'CSV row separator', "\n")
	.option('-n, --sheet-rows <num>', 'Number of rows to process (0=all rows)')
	.option('--sst', 'generate shared string table for XLS* formats')
	.option('--compress', 'use compression when writing XLSX/M/B and ODS')
	.option('--read', 'read but do not generate output')
	.option('--all', 'parse everything; write as much as possible')
	.option('--dev', 'development mode')
	.option('-q, --quiet', 'quiet mode');

program.on('--help', function() {
	console.log('  Default output format is CSV');
	console.log('  Support email: dev@sheetjs.com');
	console.log('  Website: http://sheetjs.com/');
});

/* flag, bookType, default ext */
var workbook_formats = [
	['xlsx',   'xlsx', 'xlsx'],
	['xlsm',   'xlsm', 'xlsm'],
	['xlsb',   'xlsb', 'xlsb'],
	['xls',     'xls',  'xls'],
	['biff5', 'biff5',  'xls'],
	['ods',     'ods',  'ods'],
	['fods',   'fods', 'fods']
];
var wb_formats_2 = [
	['xlml',   'xlml', 'xls']
];
program.parse(process.argv);
program.eth = program.socialcalc;

var filename = '', sheetname = '';
if(program.args[0]) {
	filename = program.args[0];
	if(program.args[1]) sheetname = program.args[1];
}
if(program.sheet) sheetname = program.sheet;
if(program.file) filename = program.file;

if(!process.stdin.isTTY) filename = filename || "-";

if(!filename) {
	console.error(n + ": must specify a filename");
	process.exit(1);
}
if(filename !== "-" && !fs.existsSync(filename)) {
	console.error(n + ": " + filename + ": No such file or directory");
	process.exit(2);
}

var opts = {};
if(program.listSheets) opts.bookSheets = true;
if(program.sheetRows) opts.sheetRows = program.sheetRows;
if(program.password) opts.password = program.password;
var seen = false;
function wb_fmt() {
	seen = true;
	opts.cellFormula = true;
	opts.cellNF = true;
	if(program.output) sheetname = program.output;
}
function isfmt(m/*:string*/)/*:boolean*/ {
	if(!program.output) return false;
	var t = m.charAt(0) === "." ? m : "." + m;
	return program.output.slice(-t.length) === t;
}
workbook_formats.forEach(function(m) { if(program[m[0]] || isfmt(m[0])) { wb_fmt(); } });
wb_formats_2.forEach(function(m) { if(program[m[0]] || isfmt(m[0])) { wb_fmt(); } });
if(seen) {
} else if(program.formulae) opts.cellFormula = true;
else opts.cellFormula = false;

var wopts = ({WTF:opts.WTF, bookSST:program.sst}/*:any*/);
if(program.compress) wopts.compression = true;

if(program.all) {
	opts.cellFormula = true;
	opts.bookVBA = true;
	opts.cellNF = true;
	opts.cellHTML = true;
	opts.cellStyles = true;
	opts.sheetStubs = true;
	opts.cellDates = true;
	wopts.cellStyles = true;
	wopts.bookVBA = true;
}
opts.dense = false;

if(program.dev) {
	opts.WTF = true;
}

var w;
function _err(e) {
	if(program.dev) throw e;
	var msg = (program.quiet) ? "" : n + ": error parsing ";
	msg += filename + ": " + e;
	console.error(msg);
	process.exit(3);
}
if(filename === "-") {
	var concat = require('concat-stream');
	process.stdin.pipe(concat(function(data) {
		try { w = X.read(data, opts); } catch(e) { _err(e); }
		process_data(w);
	}));
} else {
	try { w = X.readFile(filename, opts); } catch(e) { _err(e); }
	process_data(w);
}


function process_data(w/*:JWorkbook*/) {
var wb = w[1];
if(program.read) process.exit(0);
if(!wb) { console.error(n + ": error parsing " + filename + ": empty workbook"); process.exit(0); }

if(program.listSheets) {
	console.log((wb.SheetNames||[]).join("\n"));
	process.exit(0);
}

/* full workbook formats */
workbook_formats.forEach(function(m) { if(program[m[0]] || isfmt(m[0])) {
		wopts.bookType = m[1];
		fs.writeFileSync(sheetname || ((filename || "") + "." + m[2]), X.utils["to_" + m[1]](w, wopts));
		process.exit(0);
} });

wb_formats_2.forEach(function(m) { if(program[m[0]] || isfmt(m[0])) {
		wopts.bookType = m[1];
		fs.writeFileSync(sheetname || ((filename || "") + "." + m[2]), X.utils["to_" + m[1]](w, wopts));
		process.exit(0);
} });

var target_sheet = sheetname || '';
if(target_sheet === '') {
	if(program.sheetIndex < (wb.SheetNames||[]).length) target_sheet = wb.SheetNames[program.sheetIndex];
	else target_sheet = (wb.SheetNames||[""])[0];
}

var ws;
try {
	ws = wb.Sheets[target_sheet];
	if(!ws) {
		console.error("Sheet " + target_sheet + " cannot be found");
		process.exit(3);
	}
} catch(e) {
	console.error(n + ": error parsing "+filename+" "+target_sheet+": " + e);
	process.exit(4);
}

if(program.perf) process.exit(0);

/* single worksheet file formats */
[
	['biff2', '.xls'],
	['biff3', '.xls'],
	['biff4', '.xls'],
	//['sylk', '.slk'],
	//['html', '.html'],
	['prn', '.prn'],
	//['eth', '.eth'],
	['rtf', '.rtf'],
	['txt', '.txt'],
	['dbf', '.dbf']
	//['dif', '.dif']
].forEach(function(m) { if(program[m[0]] || isfmt(m[1])) {
	wopts.bookType = m[0];
	console.log(wopts.bookType);
	var out = X.utils["to_" + m[0]](w, wopts);
	var tgt = sheetname || w[1].SheetNames[0];
	fs.writeFileSync(sheetname || ((filename || "") + "." + m[0]), out[tgt] || out);
	process.exit(0);
} });

var oo = "";
if(!program.quiet) console.error(target_sheet);
if(program.formulae) oo = X.utils.to_formulae(w)[target_sheet].join("\n");
else if(program.json) oo = JSON.stringify(X.utils.to_json(w)[target_sheet]);
else if(program.rawJs) oo = JSON.stringify(X.utils.to_json(w,true)[target_sheet]);
else if(program.arrays) oo = JSON.stringify(X.utils.to_json(w,{raw:true, header:1})[target_sheet]);
else if(program.xml) oo = X.utils.to_xml(w)[target_sheet];
else if(program.html) oo = X.utils.to_html(w)[target_sheet];
else if(program.markdown) oo = X.utils.to_md(w)[target_sheet];
else if(program.dif) oo = X.utils.to_dif(w)[target_sheet];
else if(program.sylk) oo = X.utils.to_sylk(w)[target_sheet];
else if(program.socialcalc) oo = X.utils.to_socialcalc(w)[target_sheet];
else oo = X.utils.to_dsv(w, program.fieldSep, program.rowSep)[target_sheet];

if(program.output) fs.writeFileSync(program.output, oo);
else console.log(oo);
}
