/* j -- (C) 2013-2014 SheetJS -- http://sheetjs.com */
/* vim: set ts=2: */
var J;
var fs = require('fs'), assert = require('assert');
describe('source',function(){it('should load',function(){J=require('./');});});

var opts = {};
if(process.env.WTF) opts.WTF = true;
var ex = [".xls",".xml",".xlsx",".xlsm",".xlsb"];
if(process.env.FMTS) ex=process.env.FMTS.split(":").map(function(x){return x[0]==="."?x:"."+x;});
var exp = ex.map(function(x){ return x + ".pending"; });
function test_file(x) {	return ex.indexOf(x.substr(-4))>=0 || ex.indexOf(x.substr(-5))>=0 || exp.indexOf(x.substr(-12))>=0 || exp.indexOf(x.substr(-13))>=0; }

var files = (fs.existsSync('tests.lst') ? fs.readFileSync('tests.lst', 'utf-8').split("\n") : fs.readdirSync('test_files')).filter(test_file);

var dir = "./test_files/";

describe('should parse test files', function() {
	files.forEach(function(x) {
		it(x, x.substr(-8) == ".pending" ? null : function() {
			var wb = J.readFile(dir + x, opts);
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
