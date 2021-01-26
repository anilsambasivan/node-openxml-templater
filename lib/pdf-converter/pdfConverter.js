const fs = require('fs');
const path = require('path');
const tmp = require('tmp');
const { execFileSync} = require('child_process');

const convertToPdfSync = (document, format, filter, options) => {
    const tmpOptions = (options || {}).tmpOptions || {};
    const asyncOptions = (options || {}).asyncOptions || {};
    const tempDir = tmp.dirSync({prefix: 'libreofficeConvert_', unsafeCleanup: true, ...tmpOptions});
    const installDir = tmp.dirSync({prefix: 'soffice', unsafeCleanup: true, ...tmpOptions});
    let paths = [];
    switch (process.platform) {
        case 'darwin': 
            paths = ['/Applications/LibreOffice.app/Contents/MacOS/soffice'];
            break;
        case 'linux': 
            paths = ['/usr/bin/libreoffice', '/usr/bin/soffice'];
            break;
        case 'win32': 
            paths = [ path.join(process.env['PROGRAMFILES(X86)'], 'LIBREO~1/program/soffice.exe'),
            path.join(process.env['PROGRAMFILES(X86)'], 'LibreOffice/program/soffice.exe'),
            path.join(process.env.PROGRAMFILES, 'LibreOffice/program/soffice.exe')];
            break;
        default:
            return new Error(`Operating system not yet supported: ${process.platform}`);
    }
    if(path) {
        if (!paths.length) {
            return new Error('Could not find soffice binary');
        }
        fs.accessSync(paths[0]);
        
        fs.writeFileSync(path.join(tempDir.name, 'source'), document);

        let command = `-env:UserInstallation=file://${installDir.name} --headless --convert-to ${format}`;
        if (filter !== undefined) {
            command += `:"${filter}"`;
        }
        command += ` --outdir ${tempDir.name} ${path.join(tempDir.name, 'source')}`;
        const args = command.split(' ');
        execFileSync(paths[0], args);
        return fs.readFileSync(path.join(tempDir.name, `source.${format}`));;
    }
};

module.exports = {
    convertToPdfSync
};