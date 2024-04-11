import {Sheet} from "xlsx";

export const getMergesValue = (sheet: Sheet, x: number, y: number) => {
  const ms = sheet.merges
  for (const e of ms) {
    if (x >= e.s.c && x < e.e.c && y >= e.s.r && y < e.e.r) {
      return sheet[e.s.r]?.[e.s.c];
    }
  }
  return;
}
