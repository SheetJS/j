/* j -- (C) 2013-2014 SheetJS -- http://sheetjs.com */
/* vim: set ts=2: */
var J;
var fs = require('fs'), assert = require('assert');
describe('source',function(){it('should load',function(){J=require('./');});});

var opts = {cellNF:true};
if(process.env.WTF) opts.WTF = true;
var ex = [".xls",".xml",".xlsx",".xlsm",".xlsb"];
if(process.env.FMTS) ex=process.env.FMTS.split(":").map(function(x){return x[0]==="."?x:"."+x;});
var exp = ex.map(function(x){ return x + ".pending"; });
function test_file(x) {	return ex.indexOf(x.substr(-4))>=0 || ex.indexOf(x.substr(-5))>=0 || exp.indexOf(x.substr(-12))>=0 || exp.indexOf(x.substr(-13))>=0; }

var files = (fs.existsSync('tests.lst') ? fs.readFileSync('tests.lst', 'utf-8').split("\n") : fs.readdirSync('test_files')).filter(test_file);

var dir = "./test_files/";

files.forEach(function(x) {
	if(fs.existsSync(dir + x.replace(/\.(pending|nowrite)/, ""))) describe(x.replace(/\.pending/,""), function() {
		var wb, wbxlsx, wbxlsm, wbxlsb;
		before(function() { if(x.substr(-8) !== ".pending") wb = J.readFile(dir + x, opts); });
		it('should parse', x.substr(-8) == ".pending" ? null : function() {});
		
		it('should round-trip XLSX', x.substr(-8) == ".pending" || x.substr(-8) == ".nowrite" ? null : function() {
			fs.writeFileSync(dir + x + "__.xlsx", J.utils.to_xlsx(wb, {}));
			wbxlsx = J.readFile(dir + x + "__.xlsx", opts);
		});

		it('should round-trip XLSM', x.substr(-8) == ".pending" || x.substr(-8) == ".nowrite"  ? null : function() {
			fs.writeFileSync(dir + x + "__.xlsm", J.utils.to_xlsm(wb, {}));
			wbxlsm = J.readFile(dir + x + "__.xlsm", opts);
		});

		it.skip('should round-trip XLSB', x.substr(-8) == ".pending" || x.substr(-8) == ".nowrite" ? null : function() {
			fs.writeFileSync(dir + x + "__.xlsb", J.utils.to_xlsb(wb, {}));
			wbxlsb = J.readFile(dir + x + "__.xlsb", opts);
		});
	});
});

function cmparr(x){ for(var i=1;i!=x.length;++i) assert.deepEqual(x[0], x[i]); }

var mfopts = opts;
var mft = fs.readFileSync('multiformat.lst','utf-8').split("\n");
mft.forEach(function(x) { if(x[0]!="#") describe('MFT ' + x, function() {
	var fil = {}, f = [], r = x.split(/\s+/);
	if(r.length < 3) return;
	it('should parse all', function() {
		for(var j = 1; j != r.length; ++j) f[j-1] = J.readFile(dir + r[0] + r[j], mfopts);
	});
	it('should have the same sheetnames', function() {
		cmparr(f.map(function(x) { return x[1].SheetNames; }));
	});
	it.skip('should have the same ranges', function() {
		f[0][1].SheetNames.forEach(function(s) {
			var ss = f.map(function(x) { return x[1].Sheets[s]; });
			console.log(ss.map(function(s) { return s['!ref']; }));
			cmparr(ss.map(function(s) { return s['!ref']; }));
		});
	});
	it('should have the same merges', function() {
		f[0][1].SheetNames.forEach(function(s) {
			var ss = f.map(function(x) { return x[1].Sheets[s]; });
			cmparr(ss.map(function(s) { return (s['!merges']||[]).map(function(y) { return J.XLS.utils.encode_range(y); }).sort(); }));
		});
	});
}); });
