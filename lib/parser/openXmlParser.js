const JSZip = require('jszip');
const DOMParser = require('xmldom').DOMParser;
const XMLSerializer = require('xmldom').XMLSerializer;
const DOMImplementation = require('xmldom').DOMImplementation;
const path = require('path');
const fs = require('fs')
const { aggregateStyles, prepareStyles, generateStyles } = require('./aggregateStyles');
const { prepareMediaFiles, copyMediaFiles } = require('./aggregateMedia');
const { aggregateRelations, generateRelations } = require('./aggregateRelations');
const { aggregateContentTypes, generateContentTypes } = require('./aggregateContentType');
const { generateNumbering, prepareNumbering, aggregateNumbering } = require('./aggregateNumbering');


const PartType =  {
    Document: "word/document.xml",
    Style: "word/styles.xml",
    Numbering: "word/numbering.xml",
    FontTable: "word/fontTable.xml",
    DocumentRelations: "word/_rels/document.xml.rels",
    NumberingRelations: "word/_rels/numbering.xml.rels",
    FontRelations: "word/_rels/fontTable.xml.rels",
};

const zip = new JSZip();

let storePlaceholderChar = false;
let placeHolders = [];
let placeHolderToMerge = '';
let nodeTextContentToRecreate = '';
let nodesToMerge = [];

const createDocxSingle = async(blob, data) => {
    const z = await zip.loadAsync(blob);
    var xmlFile = zip.files[PartType.Document];
    if(xmlFile) {
        const xml = await xmlFile.async("text");
        if(xml) {
            let xmlDocument = new DOMParser().parseFromString(xml, 'text/xml');
            
            let nodesToPush = [];
            
            if(data && data.length) {
                nodeTextContentToRecreate = '';
                placeHolders = [];
                nodeTextContentToRecreate = '';
                nodesToMerge = [];
                storePlaceholderChar = false;
                traverseNodes(xmlDocument.childNodes[2].childNodes[0].childNodes, data[0]);
            }
            
            if(data && data.length > 1) {
                for(let index = 1; index < data.length; index++) {
                    nodeTextContentToRecreate = '';
                    placeHolders = [];
                    nodeTextContentToRecreate = '';
                    nodesToMerge = [];
                    storePlaceholderChar = false;
                    var xmlDocumentNext = new DOMParser().parseFromString(xml, "text/xml");
                    traverseNodes(xmlDocumentNext.childNodes[2].childNodes[0].childNodes, data[index]);
    
                    nodesToPush.push(getPageBreakNode())
                    for(let index = 0; index < xmlDocumentNext.childNodes[2].childNodes[0].childNodes.length; index++) {
                        if(xmlDocumentNext.childNodes[2].childNodes[0].childNodes[index].nodeName !== 'w:sectPr')
                            nodesToPush.push(xmlDocumentNext.childNodes[2].childNodes[0].childNodes[index]);
                    }
                }
            }
            
            nodesToPush.map((item) => {
                xmlDocument.childNodes[2].childNodes[0].appendChild(item);
            });

            var serializedXML = new XMLSerializer().serializeToString(xmlDocument);
            zip.file(PartType.Document, serializedXML);
            return await z.generateAsync({type: 'nodebuffer'});
        }
    }
}

const traverseNodes = (cNodes, jsonData) => {
    if(cNodes && cNodes.length > 0) {
        for(let index = 0; index < cNodes.length; index++) {
            let nodeItem = cNodes[index];
            if(nodeItem.nodeName ===  'w:t')
            {
                let textContent = nodeItem.textContent;
                if(textContent) {
                    for (let i = 0; i < textContent.length; i++) {
                        if(textContent.charAt(i) === '[' && !storePlaceholderChar) {
                            storePlaceholderChar = true;
                        } else if(textContent.charAt(i) === ']' && storePlaceholderChar) {
                            if(nodesToMerge && nodesToMerge.length > 0) {
                                nodesToMerge.push(nodeItem);
                                nodeTextContentToRecreate = nodeTextContentToRecreate + textContent;
                                nodeItem.textContent = nodeTextContentToRecreate;
                                nodeTextContentToRecreate = '';
                            }
                            storePlaceholderChar = false;
                            ph = placeHolderToMerge.split('||')[0];
                            placeHolders.push(ph);
                            const replaceValue = (jsonData && jsonData[ph]) ? jsonData[ph] : '';
                            const replacedText = nodeItem.textContent.replace('[' + placeHolderToMerge + ']', replaceValue)
                            nodeItem.textContent = replacedText;
                            placeHolderToMerge = '';
                        } else {
                            if(storePlaceholderChar) {
                                placeHolderToMerge = placeHolderToMerge + textContent.charAt(i);
                            }
                        }
                    }
                    if(storePlaceholderChar) {
                        nodesToMerge.push(nodeItem);
                        nodeTextContentToRecreate = nodeTextContentToRecreate + textContent;
                        nodeItem.textContent = '';
                    }
                }
            }

            traverseNodes(nodeItem.childNodes, jsonData)
        }
    }
}

const createDocxMultiple = async (templateDataSet) => {
    let documentCollection = [];
    let nodesToPush = [];
    
    for(let template of templateDataSet) {
        const buffer = fs.readFileSync(path.join(process.env.DOCUMENTS_UPLOAD_DIR + process.env.DOCX_TEMPLATE_DIR, template.file));
        const z = await zip.loadAsync(buffer);
        var xmlFile = zip.files[PartType.Document];
        if(xmlFile) {
            const xml = await xmlFile.async("text");
            if(xml) {
                let xmlDocument = new DOMParser().parseFromString(xml, 'text/xml');
                
                nodeTextContentToRecreate = '';
                placeHolders = [];
                nodeTextContentToRecreate = '';
                nodesToMerge = [];
                storePlaceholderChar = false;
                const mapData = template.data ? template.data : {};
                traverseNodes(xmlDocument.childNodes[2].childNodes[0].childNodes, mapData);
                
                documentCollection.push(xmlDocument);
            }
        }
    }

    let firstDocument = documentCollection[0];

    for(let tIndex = 1; tIndex < documentCollection.length; tIndex++ ){
        let sourceNodes = documentCollection[tIndex].childNodes[2].childNodes[0].childNodes;
        nodesToPush.push(getPageBreakNode())
        for(let i = 0; i < sourceNodes.length; i++) {
            if(sourceNodes[i].nodeName !== 'w:sectPr'){
                nodesToPush.push(sourceNodes[i]);
            }
        }
    }

    nodesToPush.map((item) => {
        firstDocument.childNodes[2].childNodes[0].appendChild(item);
    });

    const b = fs.readFileSync(path.join(process.env.DOCUMENTS_UPLOAD_DIR + process.env.DOCX_TEMPLATE_DIR, templateDataSet[0].file));
    const z = await zip.loadAsync(b);
    var serializedXML = new XMLSerializer().serializeToString(firstDocument);
    zip.file(PartType.Document, serializedXML);
    return await z.generateAsync({type: 'nodebuffer'});
}

const getPageBreakNode = () => {
    document = new DOMImplementation().createDocument(null, null, null);
    let pageBreakElement = document.createElement('w:p')
    let runElement = document.createElement('w:r');
    let brElement = document.createElement('w:br')
    
    brElement.setAttribute('w:type', 'page');
    runElement.appendChild(brElement);
    pageBreakElement.appendChild(runElement);

    return pageBreakElement;
}

const createDocxAggregateMultiple = async(templateDataSet, options = {}) => {
    let body = [];
    let header = [];
    let footer = [];
    let Basestyle = options.style || 'source';
    let style = [];
    let numbering = [];
    let pageBreak = typeof options.pageBreak !== 'undefined' ? !! options.pageBreak : true;
    let files = [];
    let contentTypes = {};

    let media = {};
    let rel = {};
    let builder = body;
    let uniqueFiles = [...new Set(templateDataSet.map(item => item.file))];
    (uniqueFiles || []).forEach(async(file) => {
        const buffer = fs.readFileSync(path.join(process.env.DOCUMENTS_UPLOAD_DIR + process.env.DOCX_TEMPLATE_DIR, file));
        const jsz = new JSZip();
        files.push(jsz);
        await jsz.loadAsync(buffer);
    });

    let documentCollection = [];
    let nodesToPush = [];
    
    for(let template of templateDataSet) {
        const buffer = fs.readFileSync(path.join(process.env.DOCUMENTS_UPLOAD_DIR + process.env.DOCX_TEMPLATE_DIR, template.file));
        
        const z = await zip.loadAsync(buffer);
        var xmlFile = zip.files[PartType.Document];
        if(xmlFile) {
            const xml = await xmlFile.async("text");
            if(xml) {
                let xmlDocument = new DOMParser().parseFromString(xml, 'text/xml');
                
                nodeTextContentToRecreate = '';
                placeHolders = [];
                nodeTextContentToRecreate = '';
                nodesToMerge = [];
                storePlaceholderChar = false;
                const mapData = template.data ? template.data : {};
                traverseNodes(xmlDocument.childNodes[2].childNodes[0].childNodes, mapData);
                
                documentCollection.push(xmlDocument);
            }
        }
    }

    let firstDocument = documentCollection[0];

    for(let tIndex = 1; tIndex < documentCollection.length; tIndex++ ){
        let sourceNodes = documentCollection[tIndex].childNodes[2].childNodes[0].childNodes;
        nodesToPush.push(getPageBreakNode())
        for(let i = 0; i < sourceNodes.length; i++) {
            if(sourceNodes[i].nodeName !== 'w:sectPr'){
                nodesToPush.push(sourceNodes[i]);
            }
        }
    }

    nodesToPush.map((item) => {
        firstDocument.childNodes[2].childNodes[0].appendChild(item);
    });

    var serializedXML = new XMLSerializer().serializeToString(firstDocument);
    //
    await aggregateContentTypes(files, contentTypes);
    await prepareMediaFiles(files, media);
    await aggregateRelations(files, rel);

    await prepareNumbering(files);
    await aggregateNumbering(files, numbering);

    await prepareStyles(files, style);
    await aggregateStyles(files, style);

    const jsZip = files[0];

    await generateContentTypes(jsZip, contentTypes);
    await copyMediaFiles(jsZip, media, files);
    await generateRelations(jsZip, rel);
    await generateNumbering(jsZip, numbering);
    await generateStyles(jsZip, style);

    jsZip.file("word/document.xml", serializedXML);

    return await jsZip.generateAsync({type: 'nodebuffer'});
}

module.exports = {
    createDocxSingle,
    createDocxMultiple,
    createDocxAggregateMultiple,
};

