import xlsx, {ParsingOptions, Sheet} from "xlsx";
import {Request} from "remotequery-ts-common";

const logger = console;
type SRecord = Record<string, string>;
type SheetRequests = { sheetname: string; requests: Request[] }
type RValueCell = { x: number; y: number; value: string; }
type RowRecordData = { kCellMap: SRecord, vCell: RValueCell }

type TableRecords = {
  name: string;
  kCellMap: SRecord;
  topAttributes: SRecord;
  records: RowRecordData[];
}
type ROptions = { serviceIdName?: string };

export function multiDimSpreadsheet2Requests(spreadsheet: Buffer, {serviceIdName = 'serviceId'}: ROptions, opts: ParsingOptions = {}): SheetRequests[] {
  const workbook = xlsx.read(spreadsheet, {...opts, type: 'buffer'})
  const result: SheetRequests[] = [];
  for (const sheetname of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetname];
    if (!sheet['!ref']) {
      continue
    }
    const range = 'A1:' + sheet['!ref'].split(':')[1];

    const jsonData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetname], {header: 1, raw: false, range});
    (jsonData as any).merges = sheet['!merges']
    const tableRecords = spreadsheetProcessor(jsonData);

    const requests = toRequests(tableRecords, serviceIdName);
    result.push({sheetname, requests});
  }
  return result;
}


function spreadsheetProcessor(sheet: xlsx.Sheet): TableRecords[] {

  let pointerY = 0;
  let topAttributes: SRecord = {};

  const requestTables: TableRecords[] = [];
  try {
    //
    // Main Loop
    //
    while (true) {
      //
      pointerY = findNonEmptyRow(sheet, pointerY);
      if (pointerY === -1) {
        break;
      }
      const m1 = getCell(sheet, 0, pointerY);
      if ("#" === m1.value) {
        pointerY++;
        continue;
      }
      const m2 = getCell(sheet, 1, pointerY);

      if (m1.value && m2.value) {
        const res = processTopAttributes(sheet, pointerY);
        topAttributes = res.topAttributes;
        pointerY = res.pointerY
      } else if (!m1.value && m2.value) {
        const res = processTable(sheet, topAttributes, pointerY);
        requestTables.push({...res.requestTable, topAttributes})
        pointerY = res.pointerY
        topAttributes = {}
      } else {
        pointerY++
      }
    }
  } catch (e) {
    logger
        .warn(`Unexpected error while table processing tab ${sheet.name}. Last pointerY ${pointerY}.`);
  }
  return requestTables;
}

export function findEmptyColumn(sheet: Sheet, xStart: number, yStart: number, yEnd: number) {
  for (let yi = yStart; yi < yEnd; yi++) {
    xStart = Math.max(xStart, (sheet[yi] || []).length);
  }
  return xStart;
}

export function findEmptyColumn2(sheet: Sheet, xStart: number, yStart: number, yEnd: number) {
  let xi = xStart;
  while (true) {
    let empty = true;
    for (let yi = yStart; yi < yEnd; yi++) {
      const row = sheet[yi];
      if (row) {
        const c = getCell(sheet, xi, yi);
        if (c.value) {
          empty = false;
          break;
        }
      }
    }
    if (empty) {
      break;
    }
    xi++;
  }
  return xi;
}

function findNonEmptyRow(sheet: Sheet, y: number) {
  for (let i = y; sheet[i]; i++) {
    if (sheet[i] && sheet[i].length > 0) {
      return i;
    }
  }
  return -1;
}


function createRequest(sheet: Sheet, vCell: RValueCell, xNameCells: RValueCell[], yNameCells: RValueCell[]): RowRecordData {
  const kCellMap: SRecord = {};
  for (const xNameCell of xNameCells) {
    const kCell = getCell(sheet, vCell.x, xNameCell.y);
    if (kCell.value) {
      kCellMap[xNameCell.value] = kCell.value;
    }
  }
  for (const yNameCell of yNameCells) {
    const kCell = getCell(sheet, yNameCell.x, vCell.y);
    if (kCell.value) {
      kCellMap[yNameCell.value] = kCell.value;
    }
  }
  return {kCellMap, vCell};// RequestData(kCellMap, vCell);
}

function processTable(sheet: Sheet, topAttributes: SRecord, pointerY: number): {
  requestTable: TableRecords,
  pointerY: number
} {

  const requestTable: TableRecords = {name: `table-${pointerY}`, records: [], topAttributes, kCellMap: {}};


  const xNameCells: RValueCell[] = [];
  const yNameCells: RValueCell[] = [];

  const endOfTableY = findEmptyRow(sheet, pointerY);
  const endOfTableX = findEmptyColumn2(sheet, 1, pointerY, endOfTableY);
  let yStart = getCell(sheet, 1, pointerY);
  while (true) {
    if (!yStart.value) {
      break;
    }
    yNameCells.push(yStart);
    yStart = getCell(sheet, yStart.x + 1, yStart.y);
  }
  let xStart = getCell(sheet, 0, pointerY + 1);
  while (true) {
    if (!xStart.value) {
      break;
    }
    xNameCells.push(xStart);
    xStart = getCell(sheet, xStart.x, xStart.y + 1);
  }
  //
  const dataStart = getCell(sheet, yStart.x, xStart.y);
  const dataEnd = getCell(sheet, endOfTableX, endOfTableY);
  //
  //

  for (let x = dataStart.x; x < dataEnd.x; x++) {
    for (let y = dataStart.y; y < dataEnd.y; y++) {
      const vCell = getCell(sheet, x, y);
      if (vCell.value) {
        const request = createRequest(sheet, vCell, xNameCells, yNameCells);
        requestTable.records.push(request);
      }
    }
  }
  return {requestTable, pointerY: dataEnd.y + 1};
}

function findEmptyRow(sheet: Sheet, y: number): number {
  for (let iy = y + 1; iy < sheet.length; iy++) {
    if (sheet[iy].length === 0) {
      return iy;
    }
  }
  return sheet.length;
}

export function getCell(sheet: Sheet, x: number, y: number): RValueCell {
  const row = sheet[y];
  const value: string = row?.[x] ?? '';
  return {
    x, y, value
  }
}

function processTopAttributes(sheet: Sheet, pointerY: number): { topAttributes: SRecord; pointerY: number } {
  const topAttributes: SRecord = {};
  while (pointerY < sheet.length) {
    const n = getCell(sheet, 0, pointerY);
    const v = getCell(sheet, 1, pointerY);
    if (n.value && v.value) {
      topAttributes[n.value] = v.value;
      pointerY++;
    } else {
      break;
    }
  }
  return {topAttributes, pointerY};
}

function toRequests(requestTables: TableRecords[] = [], serviceIdName: string) {
  const requests: Request[] = [];
  for (const requestTable of requestTables) {
    for (const e of requestTable.records) {
      const parameters: SRecord = {
        ...requestTable.topAttributes,
        ...e.kCellMap,
        '$VALUE': e.vCell.value,
        '$X': e.vCell.x.toString(),
        '$Y': e.vCell.y.toString()
      };
      const serviceId = parameters[serviceIdName] || 'skip-because-no-serviceId'
      requests.push({
        serviceId,
        parameters
      });
    }
  }
  return requests;
}

export function spreadsheet2Requests(spreadsheet: Buffer, opts: ParsingOptions = {}): Request[] {
  const workbook = xlsx.read(spreadsheet, {...opts, type: 'buffer'})

  const requests: Request[] = [];

  for (const sheetname of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetname];
    const ref =sheet['!ref']
    if (!ref) {
      continue
    }
    const range = 'A1:' + ref.split(':')[1];

    const jsonData = xlsx.utils.sheet_to_json(sheet, {raw: false,header: 1, range, blankrows:false});

    let header: string[] = [];
    const table: any[] = (jsonData || [])
    table.forEach((row: any[], index: number) => {
      if (index === 0) {
        header = row.map((c: { toString: () => any; }) => (c?.toString() ?? ''))
      } else {
        const parameters = header.reduce<SRecord>((a, h, i) => {

          return {...a, [h]: row?.[i]?.toString() ?? ''}
        }, {})
        requests.push({serviceId: sheetname, parameters})
      }
    });

  }
  return requests;
}
