import { buffer } from 'node:stream/consumers';
import fs from 'node:fs/promises';
import type JSZip from 'jszip';
import { Readable } from 'node:stream';

import WorkBookInjector from './Workbook.injector';
import SpreadSheetInjector from './Spreadsheet.injector';

export class Document {
  constructor(private readonly _input: Buffer | Readable | string) {}

  async getZip(JSZip: JSZip) {
    const input = this._input;
    let content: Buffer;
    if (input instanceof Buffer) {
      content = input;
    } else if (input instanceof Readable) {
      content = await buffer(input);
    } else {
      content = await fs.readFile(input);
    }

    return await JSZip.loadAsync(content);
  }
}
class WorkBook extends Document {
  readonly type = 'workbook';

  async inject(url: string) {
    const document = this;
    return await new WorkBookInjector(document, url).exec();
  }
}

class SpreadSheet extends Document {
  readonly type = 'workbook';

  async inject(url: string) {
    const document = this;
    return await new SpreadSheetInjector(document, url).exec();
  }
}

export class Docx extends WorkBook {}
export class Docm extends WorkBook {}
export class Dotm extends WorkBook {}
export class Dotx extends WorkBook {}

export class Xlsx extends SpreadSheet {}
export class Xlsm extends SpreadSheet {}
export class Xltm extends SpreadSheet {}
export class Xltx extends SpreadSheet {}
