// Inject tracking pixel url in document. Support formats: .xlsx, .xlsm, .xltx. Node js realization of c# realization https://github.com/wavvs/doctrack

import xml2js from 'xml2js';
import Injector from './Injector';
import JSZip from 'jszip';
import type { Document } from './Document';

export default class SpreadSheetInjector extends Injector {
  private zip: JSZip;
  private parser: xml2js.Parser;
  private builder: xml2js.Builder;
  constructor(
    private readonly _document: Document,
    private readonly _url: string,
  ) {
    super();
  }

  async exec() {
    const document = this._document;

    const zip = (this.zip = await document.getZip(JSZip));

    this.parser = new xml2js.Parser();
    this.builder = new xml2js.Builder();

    const drawing1FilePath = 'xl/drawings/drawing1.xml';
    const hasNotDrawingForSheet1 = !Boolean(zip.file(drawing1FilePath));

    const urlRId = await this.appendBlankDrawing(drawing1FilePath);
    await this.appendRelationshipBetweenDrawingAndImage({ rId: urlRId, url: this._url });

    if (hasNotDrawingForSheet1) {
      const drawingRId = await this.appendDrawingsOnSheet();
      await Promise.all([
        this.appendRelationshipBetweenDrawingsAndSheet({ rId: drawingRId, path: drawing1FilePath }),
        this.appendDrawingToContentType(drawing1FilePath),
      ]);
    }

    // console.log(`SUCCESS: outputPath ===> ${outputPath}`);
    // return await zip.generateAsync({ type: 'nodebuffer' });
    return zip.generateNodeStream({ type: 'nodebuffer' });
  }

  private async appendBlankDrawing(drawing1FilePath: string) {
    const { zip, parser, builder } = this;

    let relsXml: string;

    if (!zip.file(drawing1FilePath)) {
      relsXml = `<?xml version="1.0" encoding="utf-8"?>
    <xdr:wsDr xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
      xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
    </xdr:wsDr>`;
    } else {
      relsXml = await zip.file(drawing1FilePath)!.async('string');
    }

    const drawingResult = await parser.parseStringPromise(relsXml);

    const rId = this.generateRId();

    // New blank draw
    const newDrawing = {
      $: { editAs: 'oneCell' },
      'xdr:from': {
        'xdr:col': '0',
        'xdr:colOff': '0',
        // 'xdr:row': '0',
        'xdr:row': '1254',
        'xdr:rowOff': '0',
      },
      'xdr:to': {
        // 'xdr:col': '3',
        'xdr:col': '0',
        'xdr:colOff': '0',
        'xdr:row': '1254',
        'xdr:rowOff': '0',
      },
      'xdr:pic': {
        'xdr:nvPicPr': {
          'xdr:cNvPr': {
            $: {
              id: '1',
              name: 'Picture 1',
            },
          },
          'xdr:cNvPicPr': {
            'a:picLocks': {
              $: {
                noChangeAspect: '1',
              },
            },
          },
        },
        'xdr:blipFill': {
          'a:blip': {
            $: {
              'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
              'r:link': rId, // This ID must match the ID in the relationships file
            },
          },
          'a:stretch': {
            'a:fillRect': {},
          },
        },
        'xdr:spPr': {
          'a:xfrm': {
            'a:off': { $: { x: '0', y: '0' } },
            // 'a:ext': { $: { cx: '1000000', cy: '1000000' } },
            'a:ext': { $: { cx: '0', cy: '0' } },
          },
          'a:prstGeom': {
            $: { prst: 'rect' },
            'a:avLst': {},
          },
        },
      },
      'xdr:clientData': {},
    };

    if (!drawingResult['xdr:wsDr']['xdr:twoCellAnchor']) {
      drawingResult['xdr:wsDr']['xdr:twoCellAnchor'] = [];
    }
    drawingResult['xdr:wsDr']['xdr:twoCellAnchor'].push(newDrawing);
    // console.log(builder.buildObject(drawingResult).replace(/standalone="yes"/, ''));
    zip.file(drawing1FilePath, builder.buildObject(drawingResult).replace(/standalone="yes"/, ''));

    return rId;
  }

  private async appendRelationshipBetweenDrawingAndImage(image: { rId: string; url: string }) {
    const { zip, parser, builder } = this;

    const filePath = 'xl/drawings/_rels/drawing1.xml.rels';
    let relsXml: string;

    if (!zip.file(filePath)) {
      relsXml =
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    } else {
      relsXml = await zip.file(filePath)!.async('string');
    }

    const relsResult = await parser.parseStringPromise(relsXml);

    // Append new relationship for the external image url
    const newRel = {
      $: {
        Id: image.rId,
        Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
        Target: image.url,
        TargetMode: 'External',
      },
    };

    if (!relsResult.Relationships) {
      relsResult.Relationships = {
        $: {
          xmlns: 'http://schemas.openxmlformats.org/package/2006/relationships',
        },
        Relationship: [],
      };
    }

    if (!relsResult.Relationships.Relationship) {
      relsResult.Relationships.Relationship = [];
    }

    relsResult.Relationships.Relationship.push(newRel);

    zip.file(filePath, builder.buildObject(relsResult));
  }

  private async appendDrawingsOnSheet() {
    const { zip, parser, builder } = this;

    const rId = this.generateRId();
    const filePath = 'xl/worksheets/sheet1.xml';

    // Check if the sheet file exists
    if (!zip.file(filePath)) {
      throw new Error('Sheet1.xml does not exist in the provided XLSX file.');
    }

    const sheetXml = await zip.file(filePath)!.async('string');
    const sheetResult = await parser.parseStringPromise(sheetXml);

    // Append new drawing elements on sheet:  <drawing r:id="rId1"/>
    const newDrawing = {
      $: {
        'r:id': rId,
      },
    };

    if (!sheetResult.worksheet.drawing) {
      sheetResult.worksheet.drawing = [];
    }
    sheetResult.worksheet.drawing.push(newDrawing);

    zip.file(filePath, builder.buildObject(sheetResult));

    return rId;
  }

  private async appendRelationshipBetweenDrawingsAndSheet(drawing: { rId: string; path: string }) {
    const { zip, parser, builder } = this;
    const filePath = 'xl/worksheets/_rels/sheet1.xml.rels';

    let relsXml: string;
    if (!zip.file(filePath)) {
      relsXml =
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    } else {
      relsXml = await zip.file(filePath)!.async('string');
    }

    const relsResult = await parser.parseStringPromise(relsXml);

    // Create relationship for drawing on sheet
    const newRel = {
      $: {
        Id: drawing.rId,
        Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
        Target: '/' + drawing.path,
      },
    };

    if (!relsResult.Relationships) {
      relsResult.Relationships = {
        Relationships: {
          $: {
            xmlns: 'http://schemas.openxmlformats.org/package/2006/relationships',
          },
          Relationship: [],
        },
      };
    }

    if (!relsResult.Relationships.Relationship) {
      relsResult.Relationships.Relationship = [];
    }

    relsResult.Relationships.Relationship.push(newRel);

    zip.file(filePath, builder.buildObject(relsResult));
  }

  private async appendDrawingToContentType(drawingFilePath: string) {
    const { zip, builder, parser } = this;

    const contentTypeXml = await zip.file('[Content_Types].xml')!.async('string');
    const contentTypeResult = await parser.parseStringPromise(contentTypeXml);

    // Append ovveride for drawing:
    // <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml" />
    contentTypeResult.Types.Override.push({
      $: {
        PartName: '/' + drawingFilePath,
        ContentType: 'application/vnd.openxmlformats-officedocument.drawing+xml',
      },
    });
    zip.file('[Content_Types].xml', builder.buildObject(contentTypeResult));
  }
}
