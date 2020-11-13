# libreoffice-convert #

A simple and fast node.js module for converting office documents to different formats.

## Dependency ##

Please install libreoffice in /Applications (Mac), with your favorite package manager (Linux), or with the msi (Windows).


## Usage example ##
```javascript
import { Converter, ConvertOptions } from './converter'
import * as fs from 'fs';
import * as path from 'path';
import * as util from 'util';

const format = '.pdf';
const inputPath = path.join(__dirname, '../resources/hello.docx');
const outputPath = path.join(__dirname, `../resources/hello${format}`);

const options: ConvertOptions = {
    format: format,
    inputPath: inputPath,
    outputPath: outputPath,
    filter: undefined,
}

Converter.convertWithOptions(options).then((result) => {
    if (!result) {
        console.error('Ooops - something went wrong');
    } else {
        const writeFile = util.promisify(fs.writeFile);
        // Here in result you have pdf file which you
        // can save or transfer in another stream
        writeFile(outputPath, result).then(() => {
            console.log('Yeah! File converted to: ' + outputPath);
        }).catch(err => {
            console.error(err);
        });
    }
});
```

