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
