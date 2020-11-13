import * as fs from 'fs';
import * as path from 'path';
import * as util from 'util';
import { DirResult, dirSync } from 'tmp';
import { execFile } from 'child_process';

export interface ConvertOptions {
    format: string;
    inputPath: string;
    outputPath: string;
    filter?: string;
}

export class Converter {
    static async convertWithOptions(options: ConvertOptions): Promise<Buffer | undefined> {
        try {
            const tempDir = dirSync({ prefix: 'libreofficeConvert_', unsafeCleanup: true });
            const installDir = dirSync({ prefix: 'soffice', unsafeCleanup: true });
            const binaryPath = this.getBinaryPath();
            const readFile = util.promisify(fs.readFile);
            this.writeFile(options.inputPath, tempDir, options.format);
            this.convert(installDir, tempDir, options.format, binaryPath, options.filter);
            return readFile(path.join(tempDir.name, `source${options.format}`));
        } catch (error) {
            console.error(error);
        }
    }

    private static getBinaryPath(): string {
        try {
            let binaryPaths: string[] = [];
            switch (process.platform) {
                case 'darwin': binaryPaths = ['/Applications/Toolz/LibreOffice.app/Contents/MacOS/soffice'];
                    break;
                case 'linux': binaryPaths = ['/usr/bin/libreoffice', '/usr/bin/soffice'];
                    break;
                case 'win32': binaryPaths = [
                    // has to be improved
                    path.join(process.env['PROGRAMFILES(X86)'] !== undefined ? process.env['PROGRAMFILES(X86)'] : '', 'LIBREO~1/program/soffice.exe'),
                    path.join(process.env['PROGRAMFILES(X86)'] !== undefined ? process.env['PROGRAMFILES(X86)'] : '', 'LibreOffice/program/soffice.exe'),
                    path.join(process.env.PROGRAMFILES !== undefined ? process.env.PROGRAMFILES : '', 'LibreOffice/program/soffice.exe'),
                ];
            }
            if (fs.existsSync(binaryPaths[0])) {
                console.log('Binary found in path: ' + binaryPaths[0])
                return binaryPaths[0];
            } else {
                throw new Error('Binary has not been found.')
            }
        } catch (error) {
            throw new Error(error);
        }
    }

    private static writeFile(inputPath: string, tempDir: DirResult, format: string) {
        try {
            if (fs.existsSync(inputPath)) {
                console.log('Document found at: ' + inputPath);
                fs.writeFileSync(path.join(tempDir.name, `source${format}`), inputPath);
            } else {
                throw new Error('Something went wrong with the document path - please double check.')
            }
        } catch (error) {
            throw new Error(error);
        }
    }

    private static convert(installDir: DirResult, tempDir: DirResult, format: string, binaryPath: string, filter: string | undefined) {
        try {
            let command = `-env:UserInstallation=file://${installDir.name} --headless --convert-to ${format}`;
            if (filter !== undefined) {
                command += `:"${filter}"`;
            }
            command += ` --outdir ${tempDir.name} ${path.join(tempDir.name, 'source')}`;
            const args = command.split(' ');
            execFile(binaryPath, args);
        } catch (error) {
            throw new Error(error);
        }
    }
}
