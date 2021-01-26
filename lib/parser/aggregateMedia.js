var XMLSerializer = require('xmldom').XMLSerializer;
var DOMParser = require('xmldom').DOMParser;


const prepareMediaFiles = async(files, media) => {
    var count = 1;
    files.forEach(async(zip, index) => {
        var medFiles = zip.folder("word/media").files;

        for (var mfile in medFiles) {
            if (/^word\/media/.test(mfile) && mfile.length > 11) {
                media[count] = {};
                media[count].oldTarget = mfile;
                media[count].newTarget = mfile.replace(/[0-9]/, '_' + count).replace('word/', "");
                media[count].fileIndex = index;
                await updateMediaRelations(zip, count, media);
                await updateMediaContent(zip, count, media);
                count++;
            }
        }
    });
};

const updateMediaRelations = async(zip, count, media) => {
    var xmlString = await zip.files["word/_rels/document.xml.rels"].async('text');
    var xml = new DOMParser().parseFromString(xmlString, 'text/xml');

    var childNodes = xml.getElementsByTagName('Relationships')[0].childNodes;
    var serializer = new XMLSerializer();

    for (var node in childNodes) {
        if (/^\d+$/.test(node) && childNodes[node].getAttribute) {
            var target = childNodes[node].getAttribute('Target');
            if ('word/' + target == media[count].oldTarget) {

                media[count].oldRelID = childNodes[node].getAttribute('Id');

                childNodes[node].setAttribute('Target', media[count].newTarget);
                childNodes[node].setAttribute('Id', media[count].oldRelID + '_' + count);
            }
        }
    }

    var startIndex = xmlString.indexOf("<Relationships");
    xmlString = xmlString.replace(xmlString.slice(startIndex), serializer.serializeToString(xml.documentElement));

    zip.file("word/_rels/document.xml.rels", xmlString);
};

const updateMediaContent = async(zip, count, media) => {
    var xmlString = await zip.files["word/document.xml"].async('text');
    var xml = new DOMParser().parseFromString(xmlString, 'text/xml');

    xmlString = xmlString.replace(new RegExp(media[count].oldRelID + '"', 'g'), media[count].oldRelID + '_' + count + '"');
    zip.file("word/document.xml", xmlString);
};

const copyMediaFiles = async(base, medias, files) => {
    for (var media in medias) {
        var content = await files[medias[media].fileIndex].file(medias[media].oldTarget).async("uint8array");

        base.file('word/' + medias[media].newTarget, content);
    }
};

module.exports = {
    prepareMediaFiles,
    updateMediaRelations,
    updateMediaContent,
    copyMediaFiles
};