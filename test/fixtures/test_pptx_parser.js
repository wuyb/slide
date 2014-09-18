var PPTXParser = require('../../lib/parsers/pptx_parser');

exports.testDummy = function(test) {
	var parser = new PPTXParser();
	parser.parse();
	test.done();
}