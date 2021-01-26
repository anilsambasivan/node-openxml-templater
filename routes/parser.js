const express = require('express');
const crypto = require('crypto');
const fs = require('fs')
const multer = require('multer');
const path = require('path');

const { fileFilter } = require('../helper/uploadHelper');
const { convertToPdfSync } = require('../lib/pdf-converter/pdfConverter')
const { createDocxAggregateMultiple } = require("../lib/parser/openXmlParser");

const router = express.Router();

const storage = multer.diskStorage({
    destination: process.env.DOCUMENTS_UPLOAD_DIR + process.env.DOCX_TEMPLATE_DIR,
    filename: function (req, file, cb) {
        cb(null, "docx-" + Date.now() + path.extname(file.originalname));
    },
});


const upload = multer({
    storage: storage,
    limits: { fileSize: process.env.MAX_FILE_UPLOAD_SIZE },
    fileFilter: fileFilter,
});

router.post("/parse",
    upload.single("docx_file"),
    async (req, res, next) => {
    try {
        let params = req.body;
        let fileName = '';
        let filePath = '';

        if (req.file) {
            const data = [];
            JSON.parse(params.mappings).map((mapping) => { 
                data.push({file: req.file.filename, data: mapping}); 
            });
            const docxFileBuffer = await createDocxAggregateMultiple(data);
            switch (params.fileType) {
                case 'docx':
                    fileName = crypto.randomBytes(16).toString("hex") + `.docx`;
                    filePath = path.join(process.env.DOCUMENTS_UPLOAD_DIR + process.env.DOCX_DIR, fileName);

                    fs.writeFileSync(filePath, docxFileBuffer);
                    break;
                case 'pdf':
                    fileName = crypto.randomBytes(16).toString("hex") + `.pdf`;
                    filePath = path.join(process.env.DOCUMENTS_UPLOAD_DIR + process.env.PDF_DIR, fileName);
                    const response = convertToPdfSync(docxFileBuffer, '.pdf');

                    fs.writeFileSync(filePath, response);
                    break;
                default:
                    break;
            }
            

            res.status(200).send(fileName);
        }
    } catch (error) {
        next(error)
    }
});

router.get("/download/:type/:file",
    async (req, res, next) => {
    try {
        const type = req.params.type;
        let file = '';

        switch (type) {
            case 'pdf':
                file = path.join(process.env.DOCUMENTS_UPLOAD_DIR + process.env.PDF_DIR, `./` + req.params.file);
                break;
            case 'docx':
                file = path.join(process.env.DOCUMENTS_UPLOAD_DIR + process.env.DOCX_DIR, `./` + req.params.file);
                break;
            default:
                throw new Error('Unexpected file type!!!, Accepted file types are docx / pdf');
                break;
        }
        const src = fs.createReadStream(file);
        res.writeHead(200, {
            'Content-Type': 'application/' + type,
            'Content-Disposition': 'attachment; filename=' + req.params.file,
            'Content-Transfer-Encoding': 'Binary'
        });
        
        src.pipe(res);
    } catch (error) {
        next(error)
    }
});

module.exports = router;

