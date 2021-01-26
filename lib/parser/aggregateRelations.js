var XMLSerializer = require('xmldom').XMLSerializer;
var DOMParser = require('xmldom').DOMParser;

const aggregateRelations = async(files, rel) => {
    files.forEach(async(zip) => {
        var xmlString = await zip.files["word/_rels/document.xml.rels"].async("text");
        var xml = new DOMParser().parseFromString(xmlString, 'text/xml');

        var childNodes = xml.getElementsByTagName('Relationships')[0].childNodes;

        for (var node in childNodes) {
            if (/^\d+$/.test(node) && childNodes[node].getAttribute) {
                var Id = childNodes[node].getAttribute('Id');
                if (!rel[Id])
                rel[Id] = childNodes[node].cloneNode();
            }
        }

    });
};

const generateRelations = async(zip, rel) => {
    var xmlString = await zip.files["word/_rels/document.xml.rels"].async("text");
    var xml = new DOMParser().parseFromString(xmlString, 'text/xml');
    var serializer = new XMLSerializer();

    var types = xml.documentElement.cloneNode();

    for (var node in rel) {
        types.appendChild(rel[node]);
    }

    var startIndex = xmlString.indexOf("<Relationships");
    xmlString = xmlString.replace(xmlString.slice(startIndex), serializer.serializeToString(types));

    zip.file("word/_rels/document.xml.rels", xmlString);
};


module.exports = {
    aggregateRelations,
    generateRelations,
};