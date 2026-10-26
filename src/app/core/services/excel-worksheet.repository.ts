import { Injectable } from '@angular/core';
import { IExcelWorksheetRepository } from '../models/excel.models';

/**
 * Single Responsibility: Encapsulates all direct Excel context and worksheet operations.
 * Open/Closed Principle: Implements IExcelWorksheetRepository; can be mocked or swapped for testing.
 */
@Injectable({
  providedIn: 'root'
})
export class ExcelWorksheetRepository implements IExcelWorksheetRepository {
  public isAvailable(): boolean {
    return typeof Office !== 'undefined' && typeof Excel !== 'undefined';
  }

  public async readSelectedRange(): Promise<any[][]> {
    if (!this.isAvailable()) {
      throw new Error('Excel context is unavailable. Please run inside Microsoft Excel.');
    }

    return Excel.run(async (context) => {
      const selectedRange = context.workbook.getSelectedRange();
      selectedRange.load(['values', 'address', 'rowCount', 'columnCount']);
      await context.sync();
      return selectedRange.values;
    });
  }

  public async writeToRange(startCell: string, values: any[][]): Promise<void> {
    if (!this.isAvailable()) {
      throw new Error('Excel context is unavailable. Please run inside Microsoft Excel.');
    }

    return Excel.run(async (context) => {
      const activeSheet = context.workbook.worksheets.getActiveWorksheet();
      const numRows = values.length;
      const numCols = values[0]?.length || 1;

      const targetRange = activeSheet.getRange(startCell).getResizedRange(numRows - 1, numCols - 1);
      targetRange.values = values;

      // Format header row if multiple rows exist
      if (numRows > 1) {
        const headerRange = activeSheet.getRange(startCell).getResizedRange(0, numCols - 1);
        headerRange.format.fill.color = '#107c41'; // Excel Brand Green
        headerRange.format.font.bold = true;
        headerRange.format.font.color = '#ffffff';
      }

      activeSheet.getUsedRange().format.autofitColumns();
      await context.sync();
    });
  }

  public async highlightRow(rowIndex: number): Promise<void> {
    if (!this.isAvailable()) return;

    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      // Highlight row (1-indexed + header offset)
      const range = sheet.getRange(`A${rowIndex + 2}:G${rowIndex + 2}`);
      range.select();
      await context.sync();
    });
  }
}
