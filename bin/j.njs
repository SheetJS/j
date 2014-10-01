#!/usr/bin/env node
/* j (C) 2013-2014 SheetJS -- http://sheetjs.com */
/* vim: set ts=2 ft=javascript: */
var J;
try { J = require('../'); } catch(e) { J = require('j'); }
var fs = require('fs'), program = require('commander');
program
	.version(J.version)
	.usage('[options] <file> [sheetname]')
	.option('-f, --file <file>', 'use specified file (- for stdin)')
	.option('-s, --sheet <sheet>', 'print specified sheet (default first sheet)')
	.option('-N, --sheet-index <idx>', 'use specified sheet index (0-based)')
	.option('-l, --list-sheets', 'list sheet names and exit')
	.option('-o, --output <file>', 'output to specified file')
	.option('-B, --xlsb', 'emit XLSB to <sheetname> or <file>.xlsb')
	.option('-M, --xlsm', 'emit XLSM to <sheetname> or <file>.xlsm')
	.option('-X, --xlsx', 'emit XLSX to <sheetname> or <file>.xlsx')
	.option('-S, --formulae', 'print formulae')
	.option('-j, --json', 'emit formatted JSON (all fields text)')
	.option('-J, --raw-js', 'emit raw JS object (raw numbers)')
	.option('-x, --xml', 'emit XML')
	.option('-H, --html', 'emit HTML')
	.option('-m, --markdown', 'emit markdown table (with pipes)')
	.option('-E, --socialcalc', 'emit socialcalc')
	.option('-F, --field-sep <sep>', 'CSV field separator', ",")
	.option('-R, --row-sep <sep>', 'CSV row separator', "\n")
	.option('-n, --sheet-rows <num>', 'Number of rows to process (0=all rows)')
	.option('--dev', 'development mode')
	.option('--read', 'read but do not print out contents')
	.option('-q, --quiet', 'quiet mode');

program.on('--help', function() {
	console.log('  Default output format is CSV');
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

if(!process.stdin.isTTY) filename = filename || "-";

if(!filename) {
	console.error("j: must specify a filename");
	process.exit(1);
}

if(filename !== "-" && !fs.existsSync(filename)) {
	console.error("j: " + filename + ": No such file or directory");
	process.exit(2);
}

var opts = {}, w, X, wb;
if(program.listSheets) opts.bookSheets = true;
if(program.sheetRows) opts.sheetRows = program.sheetRows;
if(program.xlsx || program.xlsm || program.xlsb) {
	opts.cellNF = true;
	if(program.output) sheetname = program.output;
}
if(program.dev) {
	J.XLS.verbose = J.XLSX.verbose = 2;
	opts.WTF = true;
}

if(filename === "-") {
	var concat = require('concat-stream');
	process.stdin.pipe(concat(function(data) {
		if(program.dev) {
			w = J.read(data, opts);
		}
		else try {
			w = J.read(data, opts);
		} catch(e) {
			var msg = (program.quiet) ? "" : "j: error parsing ";
			msg += filename + ": " + e;
			console.error(msg);
			process.exit(3);
		}
		process_data(w);
	}));
} else {
	if(program.dev) {
		w = J.readFile(filename, opts);
	}
	else try {
		w = J.readFile(filename, opts);
	} catch(e) {
		var msg = (program.quiet) ? "" : "j: error parsing ";
		msg += filename + ": " + e;
		console.error(msg);
		process.exit(3);
	}
	process_data(w);
}


function process_data(w) {

X = w[0]; wb = w[1];
if(program.read) process.exit(0);

if(program.listSheets) {
	console.log((wb.SheetNames||[]).join("\n"));
	process.exit(0);
}

var wopts = {WTF:opts.WTF};
wopts.bookType = program.xlsm ? "xlsm" : program.xlsb ? "xlsb" : "xlsx"
if(program.xlsx) return fs.writeFileSync(sheetname || (filename + ".xlsx"), J.utils.to_xlsx(w, wopts));
if(program.xlsm) return fs.writeFileSync(sheetname || (filename + ".xlsm"), J.utils.to_xlsm(w, wopts));
if(program.xlsb) return fs.writeFileSync(sheetname || (filename + ".xlsb"), J.utils.to_xlsb(w, wopts));

var target_sheet = sheetname || '';
if(target_sheet === '') {
	if(program.sheetIndex < (wb.SheetNames||[]).length) target_sheet = wb.SheetNames[program.sheetIndex];
	else target_sheet = (wb.SheetNames||[""])[0];
}
var ws;
try {
	ws = wb.Sheets[target_sheet];
	if(!ws) throw "Sheet " + target_sheet + " cannot be found";
} catch(e) {
	console.error("j: error parsing "+filename+" "+target_sheet+": " + e);
	process.exit(4);
}

var oo = "";
if(!program.quiet) console.error(target_sheet);
if(program.formulae) oo = J.utils.to_formulae(w)[target_sheet].join("\n");
else if(program.json) oo = JSON.stringify(J.utils.to_json(w)[target_sheet]);
else if(program.rawJs) oo = JSON.stringify(J.utils.to_json(w,true)[target_sheet]);
else if(program.xml) oo = J.utils.to_xml(w)[target_sheet];
else if(program.html) oo = J.utils.to_html(w)[target_sheet];
else if(program.markdown) oo = J.utils.to_md(w)[target_sheet];
else if(program.socialcalc) oo = J.utils.to_socialcalc(w)[target_sheet];
else oo = J.utils.to_dsv(w, program.fieldSep, program.rowSep)[target_sheet];

if(program.output) fs.writeFileSync(program.output, oo);
else console.log(oo);

}
