# Doctrack

[![Actions Status](https://github.com/dm-kamaev/doctrack/workflows/Build/badge.svg)](https://github.com/dm-kamaev/doctrack/actions) ![Coverage](https://github.com/dm-kamaev/doctrack/blob/master/coverage/badge-statements.svg)

Library for injecting [tracking (spy) pixel url](https://en.wikipedia.org/wiki/Spy_pixel) into Office Documents (Office Open XML).

Support formats:
  * docx
  * docm
  * dotm
  * dotx
  * xlsx
  * xlsm
  * xltm
  * xltx

```sh
npm i doctrack -S
```

## Example
```ts
import { DocTrack, Docx, Docm, Dotm, Dotx, Xlsx, Xlsm, Xltm, Xltx } from 'doctrack';

const inputFilePath = './input.docx';
const outputFilePath = './output.docx';
const trackingPixelUrl = 'http://localhost:5001/tracking';

// xlsx - new Xlsx, docm - new Docm and etc
const doctrack = new DocTrack(new Docx(inputFilePath), trackingPixelUrl);

// Write result document to file
await doctrack.writeResultToFile(outputFilePath);
```
After open result document (output.docx) with office editor will be requested url: `http://localhost:5001/tracking`.

## Input data
You can pass input file as buffer, stream or file:
```ts
import fsPromise from 'node:fs/promises';

// As file
const doctrack = new DocTrack(new Docx('./input.docx'), trackingPixelUrl);

// As Buffer
const fileBuffer = await fsPromise.readFile('./input.docx');
const doctrack = new DocTrack(new Docx(fileBuffer), trackingPixelUrl);

// As Stream
const fileStream = fs.createReadStream('./input.docx');
const doctrack = new DocTrack(new Docx(fileStream), trackingPixelUrl);
```

## Output data
You get result file as buffer, stream or file:
```ts
// Write result document to file
await doctrack.writeResultToFile('./output.docx');

// Write result document to buffer
const buffer = await doctrack.writeResultToBuffer();

// Write result document to stream
const stream = await doctrack.writeResultToStream();
```



