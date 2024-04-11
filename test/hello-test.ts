/* tslint:disable:no-console */
/* tslint:disable:no-unused-expression */
import { expect } from 'chai';
import * as fs from 'fs';
import * as path from 'node:path';
import { multiDimSpreadsheet2Requests } from '../src/remotequery-spreadsheet';

describe('hello-test', () => {
  let data: Buffer = Buffer.from('');

  it('open file', async () => {
    data = fs.readFileSync(path.join(__dirname, 'hello.xlsx'));
    expect(data?.length > 0).true;
  });

  it('process data', async () => {
    const res = multiDimSpreadsheet2Requests(data, { serviceIdName: '$SERVICEID' });
    expect(res.length === 1).true;
    expect(res[0].requests.length === 1).true;
    expect(res[0].requests[0].serviceId).equals('InsertHello', 'Wrong serviceId!');
    expect(res[0].requests[0].parameters.name).equals('Hello', 'Wrong name!');
    expect(res[0].requests[0].parameters.value).equals('World', 'Wrong value!');
    expect(res[0].requests[0].parameters.$VALUE).equals('ok', 'Wrong value!');
  });
});
