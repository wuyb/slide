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
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    "r": 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    "pr": 'http://schemas.openxmlformats.org/package/2006/relationships',
  }
);

var selectSingleNode = function(path, root) {
  var nodes = select(path, root);
  if (nodes && nodes.length == 1) {
    return nodes[0];
  }
  return null;
}

var PPTXParser = module.exports = function () {
};

PPTXParser.prototype.parse = function(file, success, error) {

  var presentation = { slides : [], relationships : [] };

  var parseShape = function(shapeNode) {
    var shape = { paragraphs : [] };

    // read the texts (in the structured way: body->paragraph->run->text)
    var bodyNodes = select('p:txBody', shapeNode);
    if (bodyNodes && bodyNodes.length == 1) {
      _.each(select('a:p', bodyNodes[0]), function(paragraphNode) {
        var paragraph = { runs : [] };
        _.each(select('a:r', paragraphNode), function(runNode) {
          var run = { text: select('a:t/text()', runNode).toString() };
          paragraph.runs.push(run);
        });
        shape.paragraphs.push(paragraph);
      });
    }

    // get the type of the shape
    var propertiesNode = selectSingleNode('./p:nvSpPr/p:nvPr/p:ph', shapeNode);
    if (propertiesNode) {
      shape.type = propertiesNode.getAttribute('type');
    }
    return shape;
  }

  var parsePicture = function(pictureNode) {
    var picture = {};
    var nvPropertiesNode = selectSingleNode('./p:nvPicPr/p:cNvPr', pictureNode);
    if (nvPropertiesNode) {
      picture.description = nvPropertiesNode.getAttribute('descr');
    }
    var blipNode = selectSingleNode('./p:blipFill/a:blip', pictureNode);
    if (blipNode) {
      picture.relationshipId = blipNode.getAttribute('r:embed');
    }
    return picture;
  }

  var parseSlideXml = function(content) {
    var slide = { shapes : [], pictures : [] };
    var doc = new xmldom().parseFromString(content);
    _.each(select('//p:sld/p:cSld/p:spTree/p:sp', doc), function(shapeNode) {
      slide.shapes.push(parseShape(shapeNode));
    });
    _.each(select('//p:sld/p:cSld/p:spTree/p:pic', doc), function(pictureNode) {
      slide.pictures.push(parsePicture(pictureNode));
    });
    return slide;
  }

  var parseRelationshipXml = function(content) {
    var relationships = [];
    var doc = new xmldom().parseFromString(content);
    _.each(select('//pr:Relationships/pr:Relationship', doc), function(relationshipNode) {
      relationships.push({
        id : relationshipNode.getAttribute('Id'),
        target: relationshipNode.getAttribute('Target')
      });
    });
    return relationships;
  }

  var parseZipEntry = function(entry) {
    var filename = entry.path;
    if (_(filename).startsWith('ppt/slides/slide') && _(filename).endsWith('.xml')) {
      var writableStream = new streamBuffers.WritableStreamBuffer();
      writableStream.on('close', function() {
        var slide = parseSlideXml(writableStream.getContentsAsString('utf8'));
        var index = parseInt(filename.substring('ppt/slides/slide'.length, filename.indexOf('.xml')));
        presentation.slides[index - 1] = slide;
      });
      entry.pipe(writableStream);
    } else if (_(filename).startsWith('ppt/slides/_rels/slide') && _(filename).endsWith('.xml.rels')) {
      var writableStream = new streamBuffers.WritableStreamBuffer();
      writableStream.on('close', function() {
        var relationships = parseRelationshipXml(writableStream.getContentsAsString('utf8'));
        var index = parseInt(filename.substring('ppt/slides/_rels/slide'.length, filename.indexOf('.xml.rels')));
        presentation.relationships[index - 1] = relationships;
      });
      entry.pipe(writableStream);
    }
  }

  fs.createReadStream(file).pipe(unzip.Parse())
    .on('entry', parseZipEntry)
    .on('error', error)
    .on('close', function() {
      success(presentation);
    });
}
