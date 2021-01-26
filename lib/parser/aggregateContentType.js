var XMLSerializer = require('xmldom').XMLSerializer;
var DOMParser = require('xmldom').DOMParser;


const aggregateContentTypes = async(files, contentTypes) => {
    files.forEach(async(zip) => {
        var xmlString = await zip.files["[Content_Types].xml"].async("text");
        var xml = new DOMParser().parseFromString(xmlString, 'text/xml');

        var childNodes = xml.getElementsByTagName('Types')[0].childNodes;

        for (var node in childNodes) {
            if (/^\d+$/.test(node) && childNodes[node].getAttribute) {
                var contentType = childNodes[node].getAttribute('ContentType');
                if (!contentTypes[contentType])
                    contentTypes[contentType] = childNodes[node].cloneNode();
            }
        }
    });
};

const generateContentTypes = async(zip, contentTypes) => {
    var xmlString = await zip.files["[Content_Types].xml"].async("text");
    var xml = new DOMParser().parseFromString(xmlString, 'text/xml');
    var serializer = new XMLSerializer();

    var types = xml.documentElement.cloneNode();

    for (var node in contentTypes) {
        types.appendChild(contentTypes[node]);
    }

    var startIndex = xmlString.indexOf("<Types");
    xmlString = xmlString.replace(xmlString.slice(startIndex), serializer.serializeToString(types));

    zip.file("[Content_Types].xml", xmlString);
};

module.exports = {
    aggregateContentTypes,
    generateContentTypes,
};