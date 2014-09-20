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

    var slide2 = presentation.slides[1];
    test.equals(2, slide2.shapes.length);
    test.equals('title', slide2.shapes[0].type);
    test.equals('', slide2.shapes[1].type);
    test.done();
  }, function(error) {
    test.fail();
    test.done();
  });
}

// test: single page with single picture
exports.testReadPicture = function(test) {
  var parser = new PPTXParser();
  parser.parse('../test_files/picture.pptx', function(presentation) {
    test.equals(1, presentation.slides.length);
    var slide = presentation.slides[0];
    test.equals(0, slide.shapes.length);
    test.equals(1, slide.pictures.length);
    var picture = slide.pictures[0];
    test.equals('github.png', picture.description);
    test.equals('rId2', picture.relationshipId);
    test.done();
  }, function(error) {
    test.fail();
    test.done();
  });
}