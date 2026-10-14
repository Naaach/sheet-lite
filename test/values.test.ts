import { describe, it, expect } from 'vitest';
import { Workbook } from '../src/index';
import { makeMinimalXlsx } from './fixtures/make';

describe('v0.2 — reading values', () => {
  it('lists sheets and finds them by name', () => {
    const wb = Workbook.read(makeMinimalXlsx());
    expect(wb.sheetNames).toEqual(['Sheet1']);
    expect(wb.getWorksheet('Sheet1')).toBeDefined();
    expect(wb.getWorksheet('Nope')).toBeUndefined();
  });

  it('reads string and number values with correct types', () => {
    const ws = Workbook.read(makeMinimalXlsx()).getWorksheet('Sheet1')!;
    expect(ws.getCell('A1').value).toBe('Hello');
    expect(ws.getCell('A1').type).toBe('String');
    expect(ws.getCell('B1').value).toBe(42);
    expect(ws.getCell('B1').type).toBe('Number');
    expect(ws.getCell('A2').value).toBe('World');
    expect(ws.getCell('B2').value).toBeCloseTo(3.14);
    expect(ws.getCell('C9').value).toBeNull();
    expect(ws.getCell('C9').type).toBe('Null');
  });

  it('supports integer (row, col) addressing', () => {
    const ws = Workbook.read(makeMinimalXlsx()).getWorksheet('Sheet1')!;
    expect(ws.getCell(1, 1).value).toBe('Hello');
    expect(ws.getCell(1, 2).value).toBe(42);
  });

  it('iterates rows and cells', () => {
    const ws = Workbook.read(makeMinimalXlsx()).getWorksheet('Sheet1')!;
    const seen: Array<[number, number, unknown]> = [];
    ws.eachRow((row, n) => {
      row.eachCell((cell) => seen.push([n, cell.col, cell.value]));
    });
    expect(seen).toContainEqual([1, 1, 'Hello']);
    expect(seen).toContainEqual([1, 2, 42]);
    expect(seen).toContainEqual([2, 1, 'World']);
  });
});
