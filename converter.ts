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
    debug?: boolean;
}

export class Converter {
    static async convertWithOptions(options: ConvertOptions): Promise<string> {
        try {
            const tempDir = dirSync({ prefix: 'libreofficeConvert_', unsafeCleanup: true });
            const installDir = dirSync({ prefix: 'soffice', unsafeCleanup: true });
            const binaryPath = this.getBinaryPath();
            const fileExtension = this.getFileExtension(options.inputPath);
            this.copyFile(options.inputPath, `${tempDir.name}/source.${fileExtension}`);
            await this.convert(installDir, tempDir, options, fileExtension, binaryPath);
            this.copyFile(`${tempDir.name}/source.${options.format}`, options.outputPath);
            this.manualCleanup([tempDir, installDir]);
            return options.outputPath;
        } catch (error) {
            throw new Error(error);
        }
    }

    // ----------------
    // helper functions
    // ----------------

    private static manualCleanup(tmpDirs: DirResult[]) {
        for (const tmpDir of tmpDirs) {
            tmpDir.removeCallback();
        }
    }

    private static getFileExtension(path: string): string {
        try {
            const fileExtension = path.split('.').pop();
            if (!fileExtension) throw new Error();
            return fileExtension;
        } catch (error) {
            throw new Error('File has no valid extension');
        }
    }

    private static getBinaryPath(): string {
        try {
            let binaryPaths: string[] = [];
            let binaryPath: string | undefined = undefined;
            switch (process.platform) {
                case 'darwin': binaryPaths = ['/Applications/Toolz/LibreOffice.app/Contents/MacOS/soffice'];
                    break;
                case 'linux': binaryPaths = ['/usr/bin/libreoffice', '/usr/bin/soffice'];
                    break;
                case 'win32': binaryPaths = [
                    path.join(
                        process.env['PROGRAMFILES(X86)'] !== undefined ? process.env['PROGRAMFILES(X86)'] : '',
                        'LIBREO~1/program/soffice.exe'
                    ),
                    path.join(
                        process.env['PROGRAMFILES(X86)'] !== undefined ? process.env['PROGRAMFILES(X86)'] : '',
                        'LibreOffice/program/soffice.exe'
                    ),
                    path.join(
                        process.env.PROGRAMFILES !== undefined ? process.env.PROGRAMFILES : '',
                        'LibreOffice/program/soffice.exe'
                    ),
                ];
            }
            // check validity of path
            for (const path of binaryPaths) {
                if (fs.existsSync(path)) {
                    binaryPath = path;
                }
            }
            if (binaryPath) {
                return binaryPath;
            } else {
                throw new Error('Binary has not been found.');
            }
        } catch (error) {
            throw new Error(error);
        }
    }

    private static copyFile(fromPath: string, toPath: string) {
        try {
            if (fs.existsSync(fromPath)) {
                fs.copyFileSync(fromPath, toPath);
            } else {
                throw new Error('Something went wrong - please double check the paths.');
            }
        } catch (error) {
            throw new Error(error);
        }
    }

    private static async convert(
        installDir: DirResult,
        tempDir: DirResult,
        options: ConvertOptions,
        fileExtension: string,
        binaryPath: string) {
        try {
            let fileURI: string;
            if (process.platform === 'win32') {
                fileURI = `file:///${installDir.name}`;
            } else {
                fileURI = `file://${installDir.name}`;
            }
            let command = `-env:UserInstallation=${fileURI} --headless`;
            command += ` --convert-to ${options.format} ${tempDir.name}/source.${fileExtension}`;
            command += ` --outdir ${tempDir.name}`;
            if (options.filter !== undefined) {
                command += `:"${options.filter}"`;
            }
            const args = command.split(' ');
            const execFilePromise = util.promisify(execFile);
            await execFilePromise(binaryPath, args);
        } catch (error) {
            throw new Error(error);
        }
    }
}