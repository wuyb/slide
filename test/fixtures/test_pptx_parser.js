var PPTXParser = require('../../lib/parsers/pptx_parser');

// test: title can be read and the title properties are retained
exports.testReadTitle = function(test) {
  var parser = new PPTXParser();
	parser.parse('../test_files/title.pptx', function(presentation) {
    test.equals(1, presentation.slides.length);
    var slide = presentation.slides[0];
    test.equals(2, slide.shapes.length);
    test.equals('ctrTitle', slide.shapes[0].type);
    test.equals('subTitle', slide.shapes[1].type);
    test.equals('Title 1', slide.title);
    test.equals('Subtitle 1', slide.subTitle);
    test.done();
  }, function(error) {
    test.fail();
    test.done();
  });
}

// test: regular shape text can be read
exports.testReadText = function(test) {
  var parser = new PPTXParser();
  parser.parse('../test_files/text.pptx', function(presentation) {
    test.equals(1, presentation.slides.length);
    var slide = presentation.slides[0];
    test.equals(1, slide.shapes.length);
    test.equals('Text 1', slide.shapes[0].text);
    test.done();
  }, function(error) {
    test.fail();
    test.done();
  });
}

// test: multiple pages (one title page, one regular page)
exports.testReadTextInMultiPages = function(test) {
  var parser = new PPTXParser();
  parser.parse('../test_files/text_multi_pages.pptx', function(presentation) {
    test.equals(2, presentation.slides.length);
    var slide1 = presentation.slides[0];
    test.equals(2, slide1.shapes.length);
    test.equals('ctrTitle', slide1.shapes[0].type);
    test.equals('subTitle', slide1.shapes[1].type);
    test.equals('Title 1', slide1.title);
    test.equals('Subtitle 1', slide1.subTitle);

    var slide2 = presentation.slides[1];
    test.equals(2, slide2.shapes.length);
    test.equals('title', slide2.shapes[0].type);
    test.equals('Page 1', slide2.shapes[0].text);
    test.equals('', slide2.shapes[1].type);
    test.equals('Page 1 item 1Page 1 item 2', slide2.shapes[1].text);
    test.done();
  }, function(error) {
    test.fail();
    test.done();
  });
}