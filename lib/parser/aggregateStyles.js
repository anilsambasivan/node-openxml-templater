const XMLSerializer = require('xmldom').XMLSerializer;
const DOMParser = require('xmldom').DOMParser;

const prepareStyles = async(files, style) => {
    var serializer = new XMLSerializer();

    files.forEach(async(zip, index) => {
        var xmlString = await zip.files['word/styles.xml'].async("text");
        var xml = new DOMParser().parseFromString(xmlString, 'text/xml');
        var nodes = xml.getElementsByTagName('w:style');

        for (var node in nodes) {
            if (/^\d+$/.test(node) && nodes[node].getAttribute) {
                var styleId = nodes[node].getAttribute('w:styleId');
                nodes[node].setAttribute('w:styleId', styleId + '_' + index);
                var basedonStyle = nodes[node].getElementsByTagName('w:basedOn')[0];
                if (basedonStyle) {
                    var basedonStyleId = basedonStyle.getAttribute('w:val');
                    basedonStyle.setAttribute('w:val', basedonStyleId + '_' + index);
                }

                var w_next = nodes[node].getElementsByTagName('w:next')[0];
                if (w_next) {
                    var w_next_ID = w_next.getAttribute('w:val');
                    w_next.setAttribute('w:val', w_next_ID + '_' + index);
                }

                var w_link = nodes[node].getElementsByTagName('w:link')[0];
                if (w_link) {
                    var w_link_ID = w_link.getAttribute('w:val');
                    w_link.setAttribute('w:val', w_link_ID + '_' + index);
                }

                var numId = nodes[node].getElementsByTagName('w:numId')[0];
                if (numId) {
                    var numId_ID = numId.getAttribute('w:val');
                    numId.setAttribute('w:val', numId_ID + index);
                }

                await updateStyleRelContent(zip, index, styleId);
            }
        }

        var startIndex = xmlString.indexOf("<w:styles ");
        xmlString = xmlString.replace(xmlString.slice(startIndex), serializer.serializeToString(xml.documentElement));

        zip.file("word/styles.xml", xmlString);
    });
};

const aggregateStyles = async(files, styles) => {
    files.forEach(async(zip) => {
        var xml = await zip.files["word/styles.xml"].async('text');
        xml = xml.substring(xml.indexOf("<w:style "), xml.indexOf("</w:styles"));
        styles.push(xml);

    });
};

const updateStyleRelContent = async(zip, fileIndex, styleId) => {
    var xmlString = await zip.files['word/document.xml'].async('text');
    var xml = new DOMParser().parseFromString(xmlString, 'text/xml');
    xmlString = xmlString.replace(new RegExp('w:val="' + styleId + '"', 'g'), 'w:val="' + styleId + '_' + fileIndex + '"');
    zip.file("word/document.xml", xmlString);
};

const generateStyles = async(zip, styles) => {
    var xml = await zip.files["word/styles.xml"].async('text');
    var startIndex = xml.indexOf("<w:style ");
    var endIndex = xml.indexOf("</w:styles>");

    xml = xml.replace(xml.slice(startIndex, endIndex), styles.join(''));

    zip.file("word/styles.xml", xml);
};

module.exports = {
    aggregateStyles,
    prepareStyles,
    updateStyleRelContent,
    generateStyles
};