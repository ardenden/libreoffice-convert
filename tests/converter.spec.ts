import { Converter, ConvertOptions } from '../converter';
import * as path from 'path';
import * as fs from 'fs';

describe('Converter', () => {

    // global config
    const format = '.pdf';
    const inputPath = path.join(__dirname, './resources/hello.docx');
    const outputPath = path.join(__dirname, `./resources/hello${format}`);

    afterAll(() => {
        // fs.unlinkSync(outputPath + 'test');
    });

    it('should convert a docx to a pdf', () => {
        // this takes a while
        jest.setTimeout(30000)
        const options: ConvertOptions = {
            format: format,
            inputPath: inputPath,
            outputPath: outputPath,
            filter: undefined
        }
        const supportedPlatforms = ['darwin', 'linux', 'win32'];
        if (supportedPlatforms.includes(process.platform)) {
            return Converter.convertWithOptions(options).then(result => {
                expect(typeof result).toBe('string');
            });
        } else {
            return Converter.convertWithOptions(options).catch(error => {
                expect.assertions(1);
                expect(error).toMatch('error');
            });
        }
    });

    // it('if an another instance of soffice exists, should convert a word document to text',  (done) => {
    //     exec("soffice  --headless")
    //     // this command create an instance of soffice. This instance will get a failure "Error: source file could not be loaded"
    //     // but only after we ask a new convert. So this is enought to reproduce fail when an another instance is open
    //     setTimeout(()=> {
    //         const docx = _fs.readFileSync(_path.join(__dirname, '/resources/hello.docx'));
    //         convert(docx, 'txt', undefined, expectHello(done));
    //     }, 100);
    // });

});
