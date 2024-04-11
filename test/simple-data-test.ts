/* tslint:disable:no-console */
/* tslint:disable:no-unused-expression */
import { expect } from 'chai';
import * as fs from 'fs';
import * as path from 'node:path';
import { multiDimSpreadsheet2Requests } from '../src/remotequery-spreadsheet';

const filename = 'simple-data.xlsx';

describe(`processing file: ${filename} in ${__dirname}`, () => {
  let data: Buffer = Buffer.from('');

  it(`open file ${filename}`, async () => {
    data = fs.readFileSync(path.join(__dirname, filename));
    expect(data?.length > 0).true;
  });

  it(`process data for ${filename}`, async () => {
    const res = multiDimSpreadsheet2Requests(data, { serviceIdName: '$SERVICEID' });
    expect(res.length).equals(2, 'Expected sheet do not match');
    const requestsList1 = res[0];
    expect(requestsList1.sheetname).equals('tab-1', 'Wrong sheetname!');
    expect(requestsList1.requests.length).equals(6, 'Number of table request list is wrong!');

    let r = requestsList1.requests[0];
    expect(r.serviceId).equals('UserData.insert', 'serviceId is wrong!');
    expect(r.parameters.userTid).equals('11');
    expect(r.parameters.lastName).equals('Müller');
    expect(r.parameters.firstName).equals('Franz');

    r = requestsList1.requests[1];
    expect(r.serviceId).equals('UserData.insert', 'serviceId is wrong!');

    r = requestsList1.requests[2];
    expect(r.serviceId).equals('UserData.insert', 'serviceId is wrong!');

    r = requestsList1.requests[3];
    expect(r.serviceId).equals('UserData.delete', 'serviceId is wrong!');
    expect(r.parameters.$VALUE).equals('ok');
  });
});
