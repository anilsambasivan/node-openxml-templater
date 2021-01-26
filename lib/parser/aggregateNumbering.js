var XMLSerializer = require('xmldom').XMLSerializer;
var DOMParser = require('xmldom').DOMParser;

const prepareNumbering = async(files) => {
    var serializer = new XMLSerializer();
    files.forEach(async(zip, index) => {
        var xmlBin = zip.files['word/numbering.xml'];
        if (!xmlBin) {
            return;
        }
        var xmlString = await xmlBin.async('text');
        var xml = new DOMParser().parseFromString(xmlString, 'text/xml');
        var nodes = xml.getElementsByTagName('w:abstractNum');

        for (var node in nodes) {
            if (/^\d+$/.test(node) && nodes[node].getAttribute) {
                var absID = nodes[node].getAttribute('w:abstractNumId');
                nodes[node].setAttribute('w:abstractNumId', absID + index);
                var pStyles = nodes[node].getElementsByTagName('w:pStyle');
                for (var pStyle in pStyles) {
                    if (pStyles[pStyle].getAttribute) {
                        var pStyleId = pStyles[pStyle].getAttribute('w:val');
                        pStyles[pStyle].setAttribute('w:val', pStyleId + '_' + index);
                    }
                }
                var numStyleLinks = nodes[node].getElementsByTagName('w:numStyleLink');
                for (var numstyleLink in numStyleLinks) {
                    if (numStyleLinks[numstyleLink].getAttribute) {
                        var styleLinkId = numStyleLinks[numstyleLink].getAttribute('w:val');
                        numStyleLinks[numstyleLink].setAttribute('w:val', styleLinkId + '_' + index);
                    }
                }

                var styleLinks = nodes[node].getElementsByTagName('w:styleLink');
                for (var styleLink in styleLinks) {
                    if (styleLinks[styleLink].getAttribute) {
                        var styleLinkId = styleLinks[styleLink].getAttribute('w:val');
                        styleLinks[styleLink].setAttribute('w:val', styleLinkId + '_' + index);
                    }
                }
            }
        }

        var numNodes = xml.getElementsByTagName('w:num');

        for (var node in numNodes) {
            if (/^\d+$/.test(node) && numNodes[node].getAttribute) {
                var ID = numNodes[node].getAttribute('w:numId');
                numNodes[node].setAttribute('w:numId', ID + index);
                var absrefID = numNodes[node].getElementsByTagName('w:abstractNumId');
                for (var i in absrefID) {
                    if (absrefID[i].getAttribute) {
                        var iId = absrefID[i].getAttribute('w:val');
                        absrefID[i].setAttribute('w:val', iId + index);
                    }
                }
            }
        }

        var startIndex = xmlString.indexOf("<w:numbering ");
        xmlString = xmlString.replace(xmlString.slice(startIndex), serializer.serializeToString(xml.documentElement));

        zip.file("word/numbering.xml", xmlString);
    });
};

const aggregateNumbering = async(files, numbering) => {
    files.forEach(async(zip) => {
        var xmlBin = zip.files['word/numbering.xml'];
        if (!xmlBin) {
          return;
        }
        var xml = await xmlBin.async('text');

        xml = xml.substring(xml.indexOf("<w:abstractNum "), xml.indexOf("</w:numbering"));

        numbering.push(xml);

    });
};

const generateNumbering = async(zip, numbering) => {
    var xmlBin = zip.files['word/numbering.xml'];
    if (!xmlBin) {
      return;
    }
    var xml = await xmlBin.async('text');
    var startIndex = xml.indexOf("<w:abstractNum ");
    var endIndex = xml.indexOf("</w:numbering>");

    xml = xml.replace(xml.slice(startIndex, endIndex), numbering.join(''));

    zip.file("word/numbering.xml", xml);
};


module.exports = {
    prepareNumbering,
    aggregateNumbering,
    generateNumbering
};