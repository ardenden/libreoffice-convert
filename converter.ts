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
            console.log('\nStarted to convert file. This may take a while.');
            const tempDir = dirSync({ prefix: 'libreofficeConvert_', unsafeCleanup: true });
            const installDir = dirSync({ prefix: 'soffice', unsafeCleanup: true });
            const binaryPath = this.getBinaryPath(options);
            const fileExtension = this.getFileExtension(options.inputPath);
            this.copyFile(options.inputPath, `${tempDir.name}/source.${fileExtension}`, options);
            await this.convert(installDir, tempDir, options, fileExtension , binaryPath);
            this.copyFile(`${tempDir.name}/source.${options.format}`, options.outputPath, options);
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
            throw new Error ('File has no valid extension');
        }
    }

    private static getBinaryPath(options: ConvertOptions): string {
        try {
            let binaryPaths: string[] = [];
            let binaryPath: string | undefined = undefined;
            switch (process.platform) {
                case 'darwin': binaryPaths = ['/Applications/Toolz/LibreOffice.app/Contents/MacOS/soffice'];
                    break;
                case 'linux': binaryPaths = ['/usr/bin/libreoffice', '/usr/bin/soffice'];
                    break;
                case 'win32': binaryPaths = [
                    path.join(process.env['PROGRAMFILES(X86)'] !== undefined ? process.env['PROGRAMFILES(X86)'] : '', 'LIBREO~1/program/soffice.exe'),
                    path.join(process.env['PROGRAMFILES(X86)'] !== undefined ? process.env['PROGRAMFILES(X86)'] : '', 'LibreOffice/program/soffice.exe'),
                    path.join(process.env.PROGRAMFILES !== undefined ? process.env.PROGRAMFILES : '', 'LibreOffice/program/soffice.exe'),
                ];
            }
            // check validity of path
            for (const path of binaryPaths) {
                if (fs.existsSync(path)) {
                    if (options.debug) {
                        console.log('\n------------------------------------------------------');
                        console.log('Binary found at: ' + path);
                    }
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

    private static copyFile(fromPath: string, toPath: string, options: ConvertOptions) {
        try {
            if (fs.existsSync(fromPath)) {
                if (options.debug) {
                    console.log('Copy file from: ' + fromPath);
                    console.log('Copy file to: ' + toPath);
                }
                fs.copyFileSync(fromPath, toPath);
            } else {
                throw new Error('Something went wrong - please double check the paths.')
            }
        } catch (error) {
            throw new Error(error);
        }
    }

    private static async convert(installDir: DirResult, tempDir: DirResult, options: ConvertOptions, fileExtension: string, binaryPath: string) {
        try {
            let command = `-env:UserInstallation=file://${installDir.name} --headless`;
            command += ` --convert-to ${options.format} ${tempDir.name}/source.${fileExtension}`;
            command += ` --outdir ${tempDir.name}`;
            if (options.filter !== undefined) {
                command += `:"${options.filter}"`;
            }
            const args = command.split(' ');
            const execFilePromise = util.promisify(execFile);
            // await execFilePromise(binaryPath, args); 
            const { stdout, stderr } = await execFilePromise(binaryPath, args); 
            if (options.debug) {
                console.log('\n------------- DEBUG SOFFICE BINARY -------------------');
                console.log('command: ' + command);
                console.log('\nstdout:', stdout);
                console.log('stderr:', stderr);
                console.log('------------------------------------------------------\n');
            }
        } catch (error) {
            throw new Error(error);
        }
    }
}
