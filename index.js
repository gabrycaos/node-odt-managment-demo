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

// fs.createReadStream('mimmo.odt')
//   .pipe(Extract({ path: 'extract' }), console.log);
// fs.createReadStream('mimmo.odt')
//   .pipe(Parse())
//   .on('entry', entry => {
//     if(/Configurations2/.test(entry.path)) {
//       console.log(entry);
//     }
//     if (entry.path === 'content.xml') {
//       archive.append(entry, { name: entry.path })
//     } else if (entry.path === 'mimetype') {
//       archive.append(entry, { name: entry.path, zlib: { level: zlib.Z_NO_COMPRESSION } })
//     } else if (!entry.isDirectory) {
//       archive.append(entry, { name: entry.path })
//     }
//   });


// const newodt = fs.createWriteStream('mimmo_edited.odt');
// const archive = archiver('zip');
// fs.createReadStream('mimmo.odt')
//   .on('error', console.log)
//   .pipe(Extract({ path: dir }))
//   .on('finish', () => editContent())

// const editContent = () => {
//   console.log('edit...'.green);
//   fs.readFile(`${dir}/content.xml`, 'utf8', (err, contents) => {
//   parseString(contents, (err, json) => {
//     // const body = json['office:document-content']['office:body'][0]['office:text'][0];
//     // const edited = body['text:h'].concat({
//     //   _: 'terzo titolo di Mimmo spero sempre rosso',
//     //   $: {
//     //     'text:style-name': 'P4',
//     //     'text:outline-level': 3,
//     //   }
//     // });
//     // json['office:document-content']['office:body'][0]['office:text'][0] = edited;
//     console.log(json);
//     const xml = builder.buildObject(json);
//     fs.unlink(`${dir}/content.xml`, () => {
//       fs.appendFileSync(`${dir}/content.xml`, xml);
//       const archive = archiver('zip');
//       const odt = fs.createWriteStream(__dirname + '/edited.odt');
//       archive.pipe(odt);
//       fs.readdir(dir, (err, files) => {
//         if(!err) {
//           files.map(file => {
//             if (isDirectory(`${dir}/${file}`)) {
//               console.log(file.green);
//               archive.directory(`${dir}/${file}`, file);
//             } if (file === 'mimetype') {
//               archive.append(fs.createReadStream(`${dir}/${file}`), { name: file, zlib: { level: zlib.Z_NO_COMPRESSION } });
//             } else if(!isDirectory(`${dir}/${file}`)) {
//               console.log(file.rainbow);
//               archive.append(fs.createReadStream(`${dir}/${file}`), { name: file });
//             }
//           })
//         }
//       })
//     });
//   });
// })
// }