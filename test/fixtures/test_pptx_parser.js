var PPTXParser = require('../../lib/parsers/pptx_parser');

exports.testReadTitle = function(test) {
  var parser = new PPTXParser();
	parser.parse('../test_files/title.pptx', function(presentation) {
    test.equals(1, presentation.slides.length);
    var slide = presentation.slides[0];
    test.equals(2, slide.shapes.length);
    test.equals('ctrTitle', slide.shapes[0].type);
    test.equals('subTitle', slide.shapes[1].type);
    test.done();
  }, function(error) {
    test.fail();
    test.done();
  });
}