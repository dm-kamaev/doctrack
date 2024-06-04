// Inject tracking pixel url in document. Support formats: .docx, .docm, .dotx. Node js realization of c# realization https://github.com/wavvs/doctrack

import xml2js from 'xml2js';
import Injector from './Injector';
import JSZip from 'jszip';
import type { Document } from './Document';

export default class WorkBookInjector extends Injector {
  constructor(
    private readonly _document: Document,
    private readonly _url: string,
  ) {
    super();
  }

  async exec() {
    const document = this._document;
    const parser = new xml2js.Parser();
    const builder = new xml2js.Builder();

    const zip = await document.getZip(JSZip);

    const relsPath = 'word/_rels/document.xml.rels';
    const docPath = 'word/document.xml';

    const [relsXml, docXml] = await Promise.all([zip.file(relsPath)!.async('string'), zip.file(docPath)!.async('string')]);

    const [relsObj, docObj] = await Promise.all([parser.parseStringPromise(relsXml), parser.parseStringPromise(docXml)]);

    const lastId = relsObj.Relationships.Relationship.length;
    const newRId = `rId${lastId + 1}`;
    const newRelationship = {
      $: {
        Id: newRId,
        Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
        Target: this._url,
        TargetMode: 'External',
      },
    };
    relsObj.Relationships.Relationship.push(newRelationship);

    const newRelsXml = builder.buildObject(relsObj);
    zip.file(relsPath, newRelsXml);

    // console.dir(docObj['w:document']['w:body'], { depth: 3 });

    const pictureName = this.generateRId();

    // New paragraph with blank draw
    const drawing = {
      'w:r': [
        {
          'w:drawing': [
            {
              'wp:inline': [
                {
                  $: {
                    distT: '0',
                    distB: '0',
                    distL: '0',
                    distR: '0',
                  },
                  'wp:extent': [
                    {
                      $: {
                        cx: '0',
                        cy: '0',
                      },
                    },
                  ],
                  'wp:effectExtent': [
                    {
                      $: {
                        l: '0',
                        t: '0',
                        r: '0',
                        b: '0',
                      },
                    },
                  ],
                  'wp:docPr': [
                    {
                      $: {
                        id: '1',
                        name: pictureName,
                      },
                    },
                  ],
                  'wp:cNvGraphicFramePr': [
                    {
                      'a:graphicFrameLocks': [
                        {
                          $: {
                            noChangeAspect: '1',
                            'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                          },
                        },
                      ],
                    },
                  ],
                  'a:graphic': [
                    {
                      $: {
                        'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                      },
                      'a:graphicData': [
                        {
                          $: {
                            uri: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
                          },
                          'pic:pic': [
                            {
                              $: {
                                'xmlns:pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
                              },
                              'pic:nvPicPr': [
                                {
                                  'pic:cNvPr': [
                                    {
                                      $: {
                                        id: '1',
                                        name: pictureName,
                                      },
                                    },
                                  ],
                                  'pic:cNvPicPr': [{}],
                                },
                              ],
                              'pic:blipFill': [
                                {
                                  'a:blip': [
                                    {
                                      $: {
                                        'r:link': newRId,
                                        cstate: 'print',
                                      },
                                      'a:extLst': [
                                        {
                                          'a:ext': [
                                            {
                                              $: {
                                                uri: '{28A0092B-C50C-407E-A947-70E740481C1C}',
                                              },
                                            },
                                          ],
                                        },
                                      ],
                                    },
                                  ],
                                  'a:stretch': [
                                    {
                                      'a:fillRect': [{}],
                                    },
                                  ],
                                },
                              ],
                              'pic:spPr': [
                                {
                                  'a:xfrm': [
                                    {
                                      'a:off': [
                                        {
                                          $: {
                                            x: '0',
                                            y: '0',
                                          },
                                        },
                                      ],
                                      'a:ext': [
                                        {
                                          $: {
                                            cx: '0',
                                            cy: '0',
                                          },
                                        },
                                      ],
                                    },
                                  ],
                                  'a:prstGeom': [
                                    {
                                      $: {
                                        prst: 'rect',
                                      },
                                      'a:avLst': [{}],
                                    },
                                  ],
                                },
                              ],
                            },
                          ],
                        },
                      ],
                    },
                  ],
                },
              ],
            },
          ],
        },
      ],
    };

    if (!docObj['w:document']['w:body']) {
      // console.log('Initiazation!');
      docObj['w:document']['w:body'] = [{ 'w:p': [] }];
    }

    const wBody = docObj['w:document']['w:body'];
    wBody[0]['w:p'].push(drawing);

    // console.dir(docObj['w:document']['w:body'], { depth: 3 });

    zip.file(docPath, builder.buildObject(docObj));

    // return zip.generateNodeStream({ type: 'nodebuffer' });
    return zip.generateNodeStream({ type: 'nodebuffer', streamFiles: true });
    // return await zip.generateAsync({ type: 'nodebuffer' });
    // const buffer = await zip.generateAsync({ type: 'nodebuffer' });
    // await writeFile(outputPath, buffer);

    // console.log(`SUCCESS: outputPath ===> ${outputPath}`);
  }
}
