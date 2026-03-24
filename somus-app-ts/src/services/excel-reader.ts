import * as XLSX from 'xlsx';

export interface WorkbookData {
  sheetNames: string[];
  sheets: Record<string, any[][]>;
}

/**
 * Le um arquivo Excel e retorna todas as sheets como arrays 2D
 */
export async function readExcelFile(filePath: string): Promise<WorkbookData> {
  const response = await fetch(filePath);
  const buffer = await response.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array' });

  const sheets: Record<string, any[][]> = {};
  for (const name of workbook.SheetNames) {
    sheets[name] = XLSX.utils.sheet_to_json(workbook.Sheets[name], {
      header: 1,
    }) as any[][];
  }

  return {
    sheetNames: workbook.SheetNames,
    sheets,
  };
}

/**
 * Le uma sheet especifica e retorna como array 2D
 */
export async function readSheet(
  filePath: string,
  sheetName: string
): Promise<any[][]> {
  const wb = await readExcelFile(filePath);
  if (!wb.sheets[sheetName]) {
    throw new Error(`Sheet "${sheetName}" nao encontrada`);
  }
  return wb.sheets[sheetName];
}

/**
 * Le um arquivo Excel e converte para array de objetos JSON
 * usando a primeira linha como headers
 */
export async function excelToJSON(
  filePath: string,
  sheetName?: string
): Promise<Record<string, any>[]> {
  const response = await fetch(filePath);
  const buffer = await response.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array' });

  const name = sheetName || workbook.SheetNames[0];
  const sheet = workbook.Sheets[name];
  if (!sheet) {
    throw new Error(`Sheet "${name}" nao encontrada`);
  }

  return XLSX.utils.sheet_to_json(sheet) as Record<string, any>[];
}

/**
 * Le um File object do input do browser
 */
export function readExcelFromFile(file: File): Promise<Record<string, any>[]> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet) as Record<string, any>[];
        resolve(json);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

/**
 * Le um File e retorna WorkbookData completo
 */
export function readWorkbookFromFile(file: File): Promise<WorkbookData> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheets: Record<string, any[][]> = {};
        for (const name of workbook.SheetNames) {
          sheets[name] = XLSX.utils.sheet_to_json(workbook.Sheets[name], {
            header: 1,
          }) as any[][];
        }
        resolve({ sheetNames: workbook.SheetNames, sheets });
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

/**
 * Escreve dados em um arquivo Excel e dispara download
 */
export function writeExcel(
  data: any[][],
  fileName: string,
  sheetName: string = 'Sheet1'
): void {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(data);
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, fileName);
}

/**
 * Escreve dados JSON em um arquivo Excel e dispara download
 */
export function writeExcelFromJSON(
  data: Record<string, any>[],
  fileName: string,
  sheetName: string = 'Sheet1'
): void {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(data);
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, fileName);
}
