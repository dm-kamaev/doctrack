/* eslint-disable unused-imports/no-unused-imports */
/* eslint-disable @typescript-eslint/no-unused-vars */
import { DocTrack, Docx, Xlsx } from './index';
import fs from 'node:fs';
import fsPromise from 'node:fs/promises';
import path from 'node:path';
import streamPromise from 'node:stream/promises';

void (async function main() {
  // Docx
  {
    const inputFilePath = path.join(__dirname, '../input/test2.docx');
    // Pass file as stream
    // const fileStream = fs.createReadStream(inputFilePath);
    // const doctrack = new DocTrack(new Docx(fileStream), 'http://localhost:5001/tracking');

    // Pass file as buffer
    // const fileBuffer = await fsPromise.readFile(inputFilePath);
    // const doctrack = new DocTrack(new Docx(fileBuffer), 'http://localhost:5001/tracking');

    // Pass file path
    const doctrack = new DocTrack(new Docx(inputFilePath), 'http://localhost:5001/tracking');

    const outputFilePath = path.join(__dirname, '../output/test.docx');
    // Write result document to file
    await doctrack.writeResultToFile(outputFilePath);

    // Write to file
    // const buffer = await doctrack.writeResultToBuffer();
    // await fsPromise.writeFile(outputFilePath, buffer);

    // const stream = await doctrack.writeResultToStream();
    // await streamPromise.pipeline(stream, fs.createWriteStream(outputFilePath));
  }

  // Xlsx
  {
    // const inputFilePath = path.join(__dirname, '../input/test.xlsx');
    // Pass file as stream
    // const fileStream = fs.createReadStream(inputFilePath);
    // const doctrack = new DocTrack(new Xlsx(fileStream), 'http://localhost:5001/tracking');
    // Pass file as buffer
    // const fileBuffer = await fsPromise.readFile(inputFilePath);
    // const doctrack = new DocTrack(new Xlsx(fileBuffer), 'http://localhost:5001/tracking');
    // Pass file path
    // const doctrack = new DocTrack(new Xlsx(inputFilePath), 'http://localhost:5001/tracking');
    // const outputFilePath = path.join(__dirname, '../output/test.xlsx');
    // Write result document to file
    // await doctrack.writeResultToFile(outputFilePath);
    // Write to file
    // const buffer = await doctrack.writeResultToBuffer();
    // await fsPromise.writeFile(outputFilePath, buffer);
    // const stream = await doctrack.writeResultToStream();
    // await streamPromise.pipeline(stream, fs.createWriteStream(outputFilePath));
  }
})();
