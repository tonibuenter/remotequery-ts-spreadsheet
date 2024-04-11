/* tslint:disable:no-console */
/* tslint:disable:no-unused-expression */
import { expect } from 'chai';
import * as fs from 'fs';
import * as path from 'node:path';
import { multiDimSpreadsheet2Requests } from '../src/remotequery-spreadsheet';

const filename = 'complex-data.xlsx';

describe(`processing file: ${filename} in ${__dirname}`, () => {
  let data: Buffer = Buffer.from('');

  it(`open file ${filename}`, async () => {
    data = fs.readFileSync(path.join(__dirname, filename));
    expect(data?.length > 0).true;
  });

  it(`process data for ${filename}`, async () => {
    const sheetRequests = multiDimSpreadsheet2Requests(data, { serviceIdName: '$SERVICEID' });
    expect(sheetRequests.length).equals(2, 'Expected sheet do not match');
    const sheetRequest = sheetRequests[0];
    expect(sheetRequest.sheetname).equals('iPhone Sales', 'Wrong sheetname!');
    expect(sheetRequest.requests.length).equals(48, 'Expected sheet do not match');
    const r21 = sheetRequest.requests[21];
    expect(r21.serviceId).equals('iPhone-Sales', 'Wrong serviceId');
    expect(r21.parameters.currency).equals('usd', 'Wrong currency');
    expect(r21.parameters.amountType).equals('million', 'Wrong amountType');
    expect(r21.parameters.country).equals('ch', 'Wrong country');
    expect(r21.parameters.$VALUE).equals('1.2', '$VALUE is wrong!');

    const r29 = sheetRequest.requests[29];
    expect(r29.serviceId).equals('iPhone-Sales', 'Wrong serviceId');
    expect(r29.parameters.currency).equals('usd', 'Wrong currency');
    expect(r29.parameters.amountType).equals('share', 'Wrong amountType');
    expect(r29.parameters.country).equals('ch', 'Wrong country');
    expect(r29.parameters.$VALUE).equals('2%', '$VALUE is wrong!');
  });
});
