var fs = require('fs');
var unzip = require('unzip');
var streamBuffers = require('stream-buffers');
var xpath = require('xpath');
var xmldom = require('xmldom').DOMParser;
var _ = require('underscore');
_.str = require('underscore.string');
_.mixin(_.str.exports());

var select = xpath.useNamespaces(
    {
      "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main"
    });

var PPTXParser = module.exports = function () {
};

PPTXParser.prototype.parse = function(file, success, error) {

  var presentation = { slides : [] };

  fs.createReadStream(file).pipe(unzip.Parse())
    .on('entry', function(entry){
      var filename = entry.path;
      var type = entry.type;

      if (_(filename).startsWith('ppt/slides/') && _(filename).endsWith('.xml')) {

        console.log('Parsing ' + filename + ' ...');

        var slide = { shapes : [] };

        slide.index = parseInt(filename.substring('ppt/slides/slide'.length, filename.indexOf('.xml')));

        var writableStream = new streamBuffers.WritableStreamBuffer();

        writableStream.on('close', function () {

          var str = writableStream.getContentsAsString('utf8');
          var doc = new xmldom().parseFromString(str);
          var shapeNodes = select('//p:sld/p:cSld/p:spTree/p:sp', doc);

          _.each(shapeNodes, function(shapeNode) {
            var shape = { paragraphs:[] };

            // read the texts (in the structured way: body->paragraph->run->text)
            var bodyNodes = select('p:txBody', shapeNode);
            if (bodyNodes && bodyNodes.length == 1) {
              var bodyNode = bodyNodes[0];
              var paragraphNodes = select('a:p', bodyNode);
              _.each(paragraphNodes, function(paragraphNode) {
                var paragraph = { runs : [] };
                var runNodes = select('a:r', paragraphNode);
                _.each(runNodes, function(runNode) {
                  var run = { text: select('a:t/text()', runNode).toString() };
                  paragraph.runs.push(run);
                });
                shape.paragraphs.push(paragraph);
              });
            }

            // get the type of the shape
            var propertiesNodes = select('./p:nvSpPr/p:nvPr/p:ph', shapeNode);
            if (propertiesNodes && propertiesNodes.length == 1) {
              var propertyNode = propertiesNodes[0];
              shape.type = propertyNode.getAttribute('type');
            }
            slide.shapes.push(shape);
          });

          presentation.slides.push(slide);

        });

        entry.pipe(writableStream);
      }
    })
    .on('error', function(err) {
      error(err);
    })
    .on('close', function() {
      success(presentation);
    });
}