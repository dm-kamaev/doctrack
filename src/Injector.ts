import crypto from 'node:crypto';

export default class Injector {
  protected generateRId() {
    return `R${crypto.randomUUID().replaceAll('-', '')}`;
    // return crypto.randomBytes(5).toString('hex');
  }
}
// console.log(crypto.randomBytes(5).toString('hex'));
