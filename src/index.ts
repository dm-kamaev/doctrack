import fs from 'node:fs';
import streamPromise from 'node:stream/promises';
import type { Docx, Xlsx } from './Document';

export * from './Document';

export class DocTrack {
  constructor(
    private readonly _document: Docx | Xlsx,
    private readonly _url: string,
  ) {}

  async writeResultToFile(filePath: string) {
    const stream = await this._inject();
    await streamPromise.pipeline(stream, fs.createWriteStream(filePath));
    // console.log(`SUCCESS: outputPath ===> ${filePath}`);
  }

  async writeResultToBuffer() {
    const stream = await this._inject();
    return await this._streamToBuffer(stream);
  }

  async writeResultToStream() {
    return await this._inject();
  }

  private async _inject() {
    return await this._document.inject(this._url);
  }

  private _streamToBuffer(stream: NodeJS.ReadableStream): Promise<Buffer> {
    return new Promise((resolve, reject) => {
      const _buf: Buffer[] = [];
      stream.on('data', (chunk) => _buf.push(chunk));
      stream.on('end', () => resolve(Buffer.concat(_buf)));
      stream.on('error', (err) => reject(err));
    });
  }
}
