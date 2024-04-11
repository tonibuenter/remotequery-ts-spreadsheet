/* tslint:disable:no-console */
/* tslint:disable:no-unused-expression */
import { expect } from 'chai';
import * as fs from 'fs';
import * as path from 'node:path';
import { spreadsheet2Requests } from '../src/remotequery-spreadsheet';

const filename = 'spreadsheet-data.xlsx';

describe(`processing file: ${filename} in ${__dirname}`, () => {
  let data: Buffer = Buffer.from('');

  it(`open file ${filename}`, async () => {
    data = fs.readFileSync(path.join(__dirname, filename));
    expect(data?.length > 0).true;
  });

  it(`process data for ${filename}`, async () => {
    const res = spreadsheet2Requests(data);
    expect(res.length).equals(24, 'Expected sheet do not match');

    for (const r of res) {
      expect(r.serviceId).equals('iPhone-Sales', 'Wrong serviceId');
      expect(r.parameters.reportingDate).equals('2023-12-21', 'reportingDate is wrong');
      expect(r.parameters.currency).equals('usd', 'reportingDate is wrong');
      expect(r.parameters.country).oneOf(['us', 'ch', 'de'], 'country is wrong');
      expect(r.parameters.year).oneOf(['2021', '2022', '2023'], 'year is wrong');
      expect(r.parameters.product).oneOf(['iphone10', 'iphone11', 'iphone12', 'iphone13'], 'product is wrong');
      expect(parseFloat(r.parameters.million.toString()) > 0).true;
      expect(r.parameters.share.toString().endsWith('%')).true;
    }
  });
});
