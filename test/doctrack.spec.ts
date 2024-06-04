import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import stream from 'node:stream';

import { DocTrack, Docx, Docm, Dotm, Dotx, Xlsx, Xlsm, Xltm, Xltx } from '../src/index';

describe('[DocTrack]: different input data', () => {
  afterEach(() => {
    // jest.clearAllMocks();
  });

  test.each([
    [Docx, path.join(__dirname, './input/test.docx'), path.join(__dirname, `./output/${crypto.randomUUID()}.docx`)],
    [Docm, path.join(__dirname, './input/test.docm'), path.join(__dirname, `./output/${crypto.randomUUID()}.docm`)],
    [Dotm, path.join(__dirname, './input/test.dotm'), path.join(__dirname, `./output/${crypto.randomUUID()}.dotm`)],
    [Dotx, path.join(__dirname, './input/test.dotx'), path.join(__dirname, `./output/${crypto.randomUUID()}.dotx`)],
    [Xlsx, path.join(__dirname, './input/test.xlsx'), path.join(__dirname, `./output/${crypto.randomUUID()}.xlsx`)],
    [Xlsm, path.join(__dirname, './input/test.xlsm'), path.join(__dirname, `./output/${crypto.randomUUID()}.xlsm`)],
    [Xltm, path.join(__dirname, './input/test.xltm'), path.join(__dirname, `./output/${crypto.randomUUID()}.xltm`)],
    [Xltx, path.join(__dirname, './input/test.xltx'), path.join(__dirname, `./output/${crypto.randomUUID()}.xltx`)],
  ])('%s: input file path and write to file', async (DocumentType, inputFilePath, outputFilePath) => {
    const document = new DocumentType(inputFilePath);
    const docTrack = new DocTrack(document, 'http://localhost:5001/tracking');
    await docTrack.writeResultToFile(outputFilePath);

    expect(fs.existsSync(outputFilePath)).toBe(true);
  });

  test.each([
    [Docx, fs.readFileSync(path.join(__dirname, './input/test.docx')), path.join(__dirname, `./output/${crypto.randomUUID()}.docx`)],
    [Docm, fs.readFileSync(path.join(__dirname, './input/test.docm')), path.join(__dirname, `./output/${crypto.randomUUID()}.docm`)],
    [Dotm, fs.readFileSync(path.join(__dirname, './input/test.dotm')), path.join(__dirname, `./output/${crypto.randomUUID()}.dotm`)],
    [Dotx, fs.readFileSync(path.join(__dirname, './input/test.dotx')), path.join(__dirname, `./output/${crypto.randomUUID()}.dotx`)],
    [Xlsx, fs.readFileSync(path.join(__dirname, './input/test.xlsx')), path.join(__dirname, `./output/${crypto.randomUUID()}.xlsx`)],
    [Xlsm, fs.readFileSync(path.join(__dirname, './input/test.xlsm')), path.join(__dirname, `./output/${crypto.randomUUID()}.xlsm`)],
    [Xltm, fs.readFileSync(path.join(__dirname, './input/test.xltm')), path.join(__dirname, `./output/${crypto.randomUUID()}.xltm`)],
    [Xltx, fs.readFileSync(path.join(__dirname, './input/test.xltx')), path.join(__dirname, `./output/${crypto.randomUUID()}.xltx`)],
  ])('%s: input file buffer and write to file', async (DocumentType, inputFilePath, outputFilePath) => {
    const document = new DocumentType(inputFilePath);
    const docTrack = new DocTrack(document, 'http://localhost:5001/tracking');
    await docTrack.writeResultToFile(outputFilePath);

    expect(fs.existsSync(outputFilePath)).toBe(true);
  });

  test.each([
    [Docx, fs.createReadStream(path.join(__dirname, './input/test.docx')), path.join(__dirname, `./output/${crypto.randomUUID()}.docx`)],
    [Docm, fs.createReadStream(path.join(__dirname, './input/test.docm')), path.join(__dirname, `./output/${crypto.randomUUID()}.docm`)],
    [Dotm, fs.createReadStream(path.join(__dirname, './input/test.dotm')), path.join(__dirname, `./output/${crypto.randomUUID()}.dotm`)],
    [Dotx, fs.createReadStream(path.join(__dirname, './input/test.dotx')), path.join(__dirname, `./output/${crypto.randomUUID()}.dotx`)],
    [Xlsx, fs.createReadStream(path.join(__dirname, './input/test.xlsx')), path.join(__dirname, `./output/${crypto.randomUUID()}.xlsx`)],
    [Xlsm, fs.createReadStream(path.join(__dirname, './input/test.xlsm')), path.join(__dirname, `./output/${crypto.randomUUID()}.xlsm`)],
    [Xltm, fs.createReadStream(path.join(__dirname, './input/test.xltm')), path.join(__dirname, `./output/${crypto.randomUUID()}.xltm`)],
    [Xltx, fs.createReadStream(path.join(__dirname, './input/test.xltx')), path.join(__dirname, `./output/${crypto.randomUUID()}.xltx`)],
  ])('%s: input file stream and write to file', async (DocumentType, inputFilePath, outputFilePath) => {
    const document = new DocumentType(inputFilePath);
    const docTrack = new DocTrack(document, 'http://localhost:5001/tracking');
    await docTrack.writeResultToFile(outputFilePath);

    expect(fs.existsSync(outputFilePath)).toBe(true);
  });
});

describe('[DocTrack]: different output data', () => {
  afterEach(() => {
    // jest.clearAllMocks();
  });

  test.each([
    [Docx, path.join(__dirname, './input/test.docx'), path.join(__dirname, `./output/${crypto.randomUUID()}.docx`)],
    [Docm, path.join(__dirname, './input/test.docm'), path.join(__dirname, `./output/${crypto.randomUUID()}.docm`)],
    [Dotm, path.join(__dirname, './input/test.dotm'), path.join(__dirname, `./output/${crypto.randomUUID()}.dotm`)],
    [Dotx, path.join(__dirname, './input/test.dotx'), path.join(__dirname, `./output/${crypto.randomUUID()}.dotx`)],
    [Xlsx, path.join(__dirname, './input/test.xlsx'), path.join(__dirname, `./output/${crypto.randomUUID()}.xlsx`)],
    [Xlsm, path.join(__dirname, './input/test.xlsm'), path.join(__dirname, `./output/${crypto.randomUUID()}.xlsm`)],
    [Xltm, path.join(__dirname, './input/test.xltm'), path.join(__dirname, `./output/${crypto.randomUUID()}.xltm`)],
    [Xltx, path.join(__dirname, './input/test.xltx'), path.join(__dirname, `./output/${crypto.randomUUID()}.xltx`)],
  ])('%s: input file path and write to file', async (DocumentType, inputFilePath, outputFilePath) => {
    const document = new DocumentType(inputFilePath);
    const docTrack = new DocTrack(document, 'http://localhost:5001/tracking');
    await docTrack.writeResultToFile(outputFilePath);

    expect(fs.existsSync(outputFilePath)).toBe(true);
  });

  test.each([
    [Docx, fs.readFileSync(path.join(__dirname, './input/test.docx'))],
    [Docm, fs.readFileSync(path.join(__dirname, './input/test.docm'))],
    [Dotm, fs.readFileSync(path.join(__dirname, './input/test.dotm'))],
    [Dotx, fs.readFileSync(path.join(__dirname, './input/test.dotx'))],
    [Xlsx, fs.readFileSync(path.join(__dirname, './input/test.xlsx'))],
    [Xlsm, fs.readFileSync(path.join(__dirname, './input/test.xlsm'))],
    [Xltm, fs.readFileSync(path.join(__dirname, './input/test.xltm'))],
    [Xltx, fs.readFileSync(path.join(__dirname, './input/test.xltx'))],
  ])('%s: input file buffer and write to buffer', async (DocumentType, inputFilePath) => {
    const document = new DocumentType(inputFilePath);
    const docTrack = new DocTrack(document, 'http://localhost:5001/tracking');
    const buffer = await docTrack.writeResultToBuffer();

    expect(buffer).toBeInstanceOf(Buffer);
  });

  test.each([
    [Docx, fs.createReadStream(path.join(__dirname, './input/test.docx'))],
    [Docm, fs.createReadStream(path.join(__dirname, './input/test.docm'))],
    [Dotm, fs.createReadStream(path.join(__dirname, './input/test.dotm'))],
    [Dotx, fs.createReadStream(path.join(__dirname, './input/test.dotx'))],
    [Xlsx, fs.createReadStream(path.join(__dirname, './input/test.xlsx'))],
    [Xlsm, fs.createReadStream(path.join(__dirname, './input/test.xlsm'))],
    [Xltm, fs.createReadStream(path.join(__dirname, './input/test.xltm'))],
    [Xltx, fs.createReadStream(path.join(__dirname, './input/test.xltx'))],
  ])('%s: input file stream and write to stream', async (DocumentType, inputFilePath) => {
    const document = new DocumentType(inputFilePath);
    const docTrack = new DocTrack(document, 'http://localhost:5001/tracking');
    const streamOutput = await docTrack.writeResultToStream();

    expect(stream.isReadable(streamOutput)).toBe(true);
  });
});
