import fs from 'fs';
import rimraf from 'rimraf';
import zlib from 'zlib';
import colors from 'colors';
// import { Extract } from 'unzipper';
import { parseString, Builder } from 'xml2js';
import archiver from 'archiver';
import { Sink } from 'pipette';
// import { Parse, Extract } from 'unzip-stream';
import AdmZip from 'adm-zip';

const dir = '.extract';
const builder = new Builder();
const zip = new AdmZip('mimmo.odt');
const archive = archiver('zip');
const odt = fs.createWriteStream('edited.odt');
archive.pipe(odt);

const logo = `
  _(_)(_)(_)(_)_ (_)(_)(_)(_) (_)(_)(_)(_)(_)   
(_)          (_) (_)      (_)_     (_)         
(_)          (_) (_)        (_)    (_)         
(_)          (_) (_)        (_)    (_)                         _      
(_)          (_) (_)       _(_)    (_)            ._ _  ___  _| | ___ 
(_)_  _  _  _(_) (_)_  _  (_)      (_)            | ' |/ . \/ . |/ ._>
(_)(_)(_)(_)    (_)(_)(_)(_)       (_)            |_|_|\___/\___|\___.
`.rainbow;


console.log(logo);

console.log(`${'LIBREOFFICE'.green} conference 2018`);

zip.getEntries().map(entry => {
  if (entry.entryName === 'content.xml') {
    parseString(entry.getData().toString('utf8'), (err, json) => {
      json['office:document-content']['office:body'][0]['office:text'][0]['text:h'][0]['_'] = 'Un nuovo Titolo';
      json['office:document-content']['office:body'][0]['office:text'][0]['text:h'][1] = {
        '_': 'Sottotitolo',
        $: {
          'text:style-name': 'P2',
          'text:outline-level': '1',
        }
      };
      json['office:document-content']['office:body'][0]['office:text'][0]['text:p'] = [
        {
          '_': 'Paragrafo',
          $: {
            'text:style-name': 'P2',
            'text:outline-level': '1',
          }
        }
      ]
      const xml = builder.buildObject(json);
      archive.append(Buffer.from(xml), { name: 'content.xml' });
    });
  } else if (entry.entryName === 'mimetype') {
    archive.append(Buffer.from(entry.getData().toString('utf8')), { name: 'mimetype', zlib: zlib.constants.Z_NO_COMPRESSION });
  } else {
    archive.append(Buffer.from(entry.getData().toString('utf8')), { name: entry.entryName });
  }
});

archive.finalize();
