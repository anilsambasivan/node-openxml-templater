

const fileFilter = (req, file, cb) => {
    // Accept images only
    if (!file.originalname.match(/\.(pdf|docx|PDF|DOCX)$/)) {
        req.fileValidationError = 'Only pdf/docx files are allowed to upload!';
        return cb(new Error('Only pdf/docx files are allowed to upload!'), false);
    }
    cb(null, true);
};

module.exports ={
    fileFilter: fileFilter,
}