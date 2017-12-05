/* j (C) 2013-present SheetJS -- http://sheetjs.com */
/* vim: set ts=2: */
/*jshint mocha:true */
/*global process, require */
/*::
declare type EmptyFunc = (() => void) | null;
declare type DescribeIt = { (desc:string, test:EmptyFunc):void; skip(desc:string, test:EmptyFunc):void; };
declare var describe : DescribeIt;
declare var it: DescribeIt;
declare var before:(test:EmptyFunc)=>void;
declare var cptable: any;
*/
var J;
var modp = './';
var fs = require('fs'), assert = require('assert');
describe('source',function(){it('should load',function(){J=require(modp);});});

var opts = ({cellNF: true}/*:any*/);
if(process.env.WTF) opts.WTF = true;
var ex = [".xls",".xml",".xlsx",".xlsm",".xlsb",".csv",".slk",".dif",".txt"];
if(process.env.FMTS) ex=process.env.FMTS.split(":").map(function(x){return x[0]==="."?x:"."+x;});
var exp = ex.map(function(x){ return x + ".pending"; });
function test_file(x){ return ex.indexOf(x.slice(-5))>=0||exp.indexOf(x.slice(-13))>=0 || ex.indexOf(x.slice(-4))>=0||exp.indexOf(x.slice(-12))>=0; }

var files = (fs.existsSync('tests.lst') ? fs.readFileSync('tests.lst', 'utf-8').split("\n").map(function(x) { return x.trim(); }) : fs.readdirSync('test_files')).filter(test_file);

var dir = "./test_files/";
var outdir = "./tmp/";

before(function(){if(!fs.existsSync(dir))throw new Error(dir + " missing");});

files.forEach(function(x) {
	if(fs.existsSync(dir + x.replace(/\.(pending|nowrite|noods)/, ""))) describe(x.replace(/\.pending/,""), function() {
		var _x = x.replace(/\.(pending|nowrite|noods)/, "");
		var pending = x.slice(-8) == ".pending";
		var nowrite = x.slice(-8) == ".nowrite";
		var noods = x.slice(-6) == ".noods";
		var wb;
		before(function() { if(!pending) wb = J.readFile(dir + _x, opts); });
		it('should parse', pending ? null : function() {});
		it('should generate files', pending ? null : function() {
			J.utils.to_formulae(wb);
			J.utils.to_json(wb, true);
			J.utils.to_json(wb, false);
			J.utils.to_dsv(wb,",", "\n");
			J.utils.to_dsv(wb,";", "\n");
			J.utils.to_html(wb);
			J.utils.to_html_cols(wb);
			J.utils.to_md(wb);
			J.utils.to_xml(wb);
		});

		it('should round-trip XLSX', pending || nowrite ? null : function() {
			fs.writeFileSync(outdir + _x + "__.xlsx", J.utils.to_xlsx(wb, {bookSST:true}));
			J.readFile(outdir + _x + "__.xlsx", opts);
		});

		it('should round-trip XLSB', pending || nowrite ? null : function() {
			fs.writeFileSync(outdir + _x + "__.xlsb", J.utils.to_xlsb(wb, {bookSST:true}));
			J.readFile(outdir + _x + "__.xlsb", opts);
		});

		it.skip('should round-trip ODS', pending || nowrite || noods ? null : function() {
			fs.writeFileSync(outdir + _x + "__.ods", J.utils.to_ods(wb));
			J.readFile(outdir + _x + "__.ods", opts);
		});

		it.skip('should round-trip FODS', pending || nowrite || noods ? null : function() {
			fs.writeFileSync(outdir + _x + "__.fods", J.utils.to_fods(wb));
			J.readFile(outdir + _x + "__.fods", opts);
		});

		it('should round-trip BIFF2 XLS', pending || nowrite ? null : function() {
			fs.writeFileSync(outdir + _x + "__.biff2", J.utils.to_biff2(wb));
			J.read(fs.readFileSync(outdir + _x + "__.biff2"), opts);
		});

		it.skip('should round-trip DIF', pending || nowrite ? null : function() {
			fs.writeFileSync(outdir + _x + "__.dif", J.utils.to_dif(wb)[wb[1].SheetNames[0]]);
			J.readFile(outdir + _x + "__.dif", opts);
		});

		it.skip('should round-trip SYLK', pending || nowrite ? null : function() {
			fs.writeFileSync(outdir + _x + "__.slk", J.utils.to_sylk(wb)[wb[1].SheetNames[0]]);
			J.readFile(outdir + _x + "__.slk", opts);
		});

		it('should round-trip ETH', pending || nowrite ? null : function() {
			fs.writeFileSync(outdir + _x + "__.eth", J.utils.to_socialcalc(wb)[wb[1].SheetNames[0]]);
			J.readFile(outdir + _x + "__.eth", opts);
		});
	});
});

function cmparr(x){ for(var i=1;i!=x.length;++i) assert.deepEqual(x[0], x[i]); }

describe('multiformat tests', function() {
var mfopts = opts;
var mft = fs.readFileSync('multiformat.lst','utf-8').split("\n");
var csv = true;
mft.forEach(function(x) {
	if(x.charAt(0)!="#") describe('MFT ' + x, function() {
		var f = [], r = x.split(/\s+/);
		if(r.length < 3) return;
		if(!fs.existsSync(dir + r[0] + r[1])) return;
		it('should parse all', function() {
			for(var j = 1; j < r.length; ++j) f[j-1] = J.readFile(dir + r[0] + r[j], mfopts);
		});
		it('should have the same sheetnames', function() {
			cmparr(f.map(function(x) { return x[1].SheetNames; }));
		});
		it('should have the same ranges', function() {
			f[0][1].SheetNames.forEach(function(s) {
				var ss = f.map(function(x) { return x[1].Sheets[s]; });
				cmparr(ss.map(function(s) { return s['!ref']; }));
			});
		});
		it('should have the same merges', function() {
			f[0][1].SheetNames.forEach(function(s) {
				var ss = f.map(function(x) { return x[1].Sheets[s]; });
				cmparr(ss.map(function(s) { return (s['!merges']||[]).map(function(y) { return J.XLS.utils.encode_range(y); }).sort(); }));
			});
		});
		it('should have the same CSV', csv ? function() {
			cmparr(f.map(function(x) { return x[1].SheetNames; }));
			var names = f[0][1].SheetNames;
			names.forEach(function(name) {
				cmparr(f.map(function(x) { return J.utils.to_csv(x)[name]; }));
			});
		} : null);
	});
	else x.split(/\s+/).forEach(function(w) { switch(w) {
		case "no-csv": csv = false; break;
		case "yes-csv": csv = true; break;
	}});
}); });

