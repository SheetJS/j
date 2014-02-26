#!/usr/bin/env node
/* j (C) 2013-2014 SheetJS -- http://sheetjs.com */
/* vim: set ts=2: */
var J;
try { J = require('../'); } catch(e) { J = require('j'); }
var fs = require('fs'), program = require('commander');
program
	.version(J.version)
	.usage('[options] <file> [sheetname]')
	.option('-f, --file <file>', 'use specified workbook')
	.option('-s, --sheet <sheet>', 'print specified sheet (default first sheet)')
	.option('-l, --list-sheets', 'list sheet names and exit')
	.option('-S, --formulae', 'print formulae')
	.option('-j, --json', 'emit formatted JSON rather than CSV (all fields text)')
	.option('-J, --raw-js', 'emit raw JS object rather than CSV (raw numbers)')
	.option('-X, --xml', 'emit XML rather than CSV')
	.option('-H, --html', 'emit HTML rather than CSV')
	.option('-F, --field-sep <sep>', 'CSV field separator', ",")
	.option('-R, --row-sep <sep>', 'CSV row separator', "\n")
	.option('--dev', 'development mode')
	.option('--read', 'read but do not print out contents')
	.option('-q, --quiet', 'quiet mode');

program.on('--help', function() {
	console.log('  Support email: dev@sheetjs.com');
	console.log('  Web Demo: http://oss.sheetjs.com/');
});

program.parse(process.argv);

var filename, sheetname = '';
if(program.args[0]) {
	filename = program.args[0];
	if(program.args[1]) sheetname = program.args[1];
}
if(program.sheet) sheetname = program.sheet;
if(program.file) filename = program.file;

if(!filename) {
	console.error("j: must specify a filename");
	process.exit(1);
}

if(!fs.existsSync(filename)) {
	console.error("j: " + filename + ": No such file or directory");
	process.exit(2);
}

var opts = {}, w, X, wb;
if(program.listSheets) opts.bookSheets = true;

if(program.dev) {
	J.XLS.verbose = J.XLSX.verbose = 2;
	opts.WTF = true;
	w = J.readFile(filename, opts);
	X = w[0]; wb = w[1];
}
else try {
	w = J.readFile(filename, opts);
	X = w[0]; wb = w[1];
} catch(e) {
	var msg = (program.quiet) ? "" : "j: error parsing ";
	msg += filename + ": " + e;
	console.error(msg);
	process.exit(3);
}
if(program.read) process.exit(0);

if(program.listSheets) {
	console.log((wb.SheetNames||[]).join("\n"));
	process.exit(0);
}

var target_sheet = sheetname || '';
if(target_sheet === '') target_sheet = (wb.SheetNames||[""])[0];

var ws;
try {
	ws = wb.Sheets[target_sheet];
	if(!ws) throw "Sheet " + target_sheet + " cannot be found";
} catch(e) {
	console.error("j: error parsing "+filename+" "+target_sheet+": " + e);
	process.exit(4);
}

var o;
if(!program.quiet) console.error(target_sheet);
if(program.formulae) o= J.utils.to_formulae(w)[target_sheet].join("\n");
else if(program.json) o= JSON.stringify(J.utils.to_json(w)[target_sheet]);
else if(program.rawJs) o= JSON.stringify(J.utils.to_json(w,true)[target_sheet]);
else if(program.xml) o= J.utils.to_xml(w)[target_sheet];
else if(program.html) o= J.utils.to_html(w)[target_sheet];
else o= J.utils.to_dsv(w, program.fieldSep, program.rowSep)[target_sheet];

console.log(o);
