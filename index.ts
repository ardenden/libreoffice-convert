import { Converter, ConvertOptions } from './converter'
import * as path from 'path';

const format = 'pdf';
const inputPath = path.join(__dirname, '../resources/hello.docx');
const outputPath = path.join(__dirname, `../resources/hello.${format}`);

const options: ConvertOptions = {
    format: format,
    inputPath: inputPath,
    outputPath: outputPath,
    filter: undefined,
    debug: true
}

Converter.convertWithOptions(options).then(result => {
    console.log('File converted to: ' + result);
}).catch(error => {
    console.error(error);
});