// import modules
import fs from 'fs';
import zlib from 'zlib';
import colors from 'colors';
import { parseString, Builder } from 'xml2js';
import archiver from 'archiver';
import AdmZip from 'adm-zip';
import Logo from './logo';

// create the read and write streams for read odt file and write the new
const builder = new Builder();
const zip = new AdmZip('demo.odt');
const archive = archiver('zip');
const odt = fs.createWriteStream('edited.odt');
archive.pipe(odt);

console.log(Logo());
console.log(`${'LIBREOFFICE'.green} conference 2018`);

zip.getEntries().map(entry => {
  console.log(`Parsing ${entry.entryName}...`.green);
  if (entry.entryName === 'content.xml') {
    parseString(entry.getData().toString('utf8'), (err, json) => {
      // We can use Javascript Array destructuring here
      // Changing title...
      json['office:document-content']['office:body'][0]['office:text'][0]['text:h'][0]['_'] = 'A new Title';
      // Adding a subtitle
      json['office:document-content']['office:body'][0]['office:text'][0]['text:h'][1] = {
        '_': 'Subtitle',
        $: {
          'text:style-name': 'P2',
          'text:outline-level': '1',
        }
      };
      // Adding a Paragraph
      json['office:document-content']['office:body'][0]['office:text'][0]['text:p'] = [
        {
          '_': 'Paragraph text: blah blah blah',
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
console.log('all files has been parsed, building new odt...');
archive.finalize();
console.log('new odt file has been created! ./edited.odt'.green);
