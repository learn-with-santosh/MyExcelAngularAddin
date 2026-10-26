import { Injectable } from '@angular/core';

export interface SheetRowData {
  id: number;
  product: string;
  category: string;
  units: number;
  unitPrice: number;
  revenue: number;
  date: string;
}

@Injectable({
  providedIn: 'root'
})
export class OfficeService {
  private isOfficeReady = false;

  constructor() {
    this.checkOfficeReady();
  }

  public async checkOfficeReady(): Promise<boolean> {
    if (typeof Office !== 'undefined') {
      try {
        await Office.onReady();
        this.isOfficeReady = true;
        return true;
      } catch (err) {
        console.warn('Office.js initialization warning:', err);
        this.isOfficeReady = false;
        return false;
      }
    }
    return false;
  }

  public isInitialized(): boolean {
    return this.isOfficeReady && typeof Excel !== 'undefined';
  }

  /**
   * Reads data from the active selection in Excel.
   */
  public async readSelectedRange(): Promise<any[][]> {
    if (!this.isInitialized()) {
      return Promise.reject(new Error('Excel is not running. Please open inside Excel to read worksheet range.'));
    }

    return Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(['values', 'address', 'rowCount', 'columnCount']);
      await context.sync();
      return range.values;
    });
  }

  /**
   * Writes tabular 2D array data into the active worksheet starting at specified address.
   */
  public async writeToRange(startCell: string, values: any[][]): Promise<void> {
    if (!this.isInitialized()) {
      return Promise.reject(new Error('Excel is not running. Please open inside Excel to write to worksheet.'));
    }

    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const numRows = values.length;
      const numCols = values[0]?.length || 1;
      
      const targetRange = sheet.getRange(startCell).getResizedRange(numRows - 1, numCols - 1);
      targetRange.values = values;
      
      // Auto format header if more than 1 row
      if (numRows > 1) {
        const headerRange = sheet.getRange(startCell).getResizedRange(0, numCols - 1);
        headerRange.format.fill.color = '#107c41'; // Excel green
        headerRange.format.font.bold = true;
        headerRange.format.font.color = '#ffffff';
      }

      sheet.getUsedRange().format.autofitColumns();
      await context.sync();
    });
  }

  /**
   * Highlights a specific row in the Excel worksheet
   */
  public async highlightRow(rowIndex: number): Promise<void> {
    if (!this.isInitialized()) return;

    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      // Highlight row (1-indexed + header offset)
      const range = sheet.getRange(`A${rowIndex + 2}:F${rowIndex + 2}`);
      range.select();
      await context.sync();
    });
  }

  /**
   * Inserts starter sales data into the active sheet to quickly test Kendo UI Grid & Charts
   */
  public async populateSampleSalesData(): Promise<void> {
    const data: any[][] = [
      ['ID', 'Product', 'Category', 'Units', 'Unit Price', 'Revenue', 'Date'],
      [101, 'Surface Pro 9', 'Hardware', 24, 999.99, 23999.76, '2026-03-15'],
      [102, 'Office 365 E5', 'Software', 150, 38.00, 5700.00, '2026-03-18'],
      [103, 'Azure Synapse', 'Cloud', 12, 1200.00, 14400.00, '2026-03-20'],
      [104, 'Surface Laptop Studio', 'Hardware', 18, 1599.00, 28782.00, '2026-03-22'],
      [105, 'Power BI Premium', 'Software', 45, 20.00, 900.00, '2026-03-25'],
      [106, 'Dynamics 365 CRM', 'Software', 30, 95.00, 2850.00, '2026-03-28'],
      [107, 'Azure OpenAI Service', 'Cloud', 80, 250.00, 20000.00, '2026-03-29']
    ];

    await this.writeToRange('A1', data);
  }
}
