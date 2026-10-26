import { Injectable } from '@angular/core';
import { IExcelMarketWriter, MarketQuote, HistoricalBar, MarketNewsItem } from '../models/market-data.models';

/**
 * Single Responsibility: Translates Market Data entities into styled Excel Tables and Formulas.
 */
@Injectable({
  providedIn: 'root'
})
export class ExcelMarketWriter implements IExcelMarketWriter {
  public isOfficeAvailable(): boolean {
    return typeof Office !== 'undefined' && typeof Excel !== 'undefined';
  }

  public async getSelectedCellAddress(): Promise<string> {
    if (!this.isOfficeAvailable()) return 'A2';

    return Excel.run(async (context) => {
      const selected = context.workbook.getSelectedRange();
      selected.load('address');
      await context.sync();
      return selected.address.split('!').pop() || 'A2';
    });
  }

  public async insertFormula(cellAddress: string, formula: string): Promise<void> {
    if (!this.isOfficeAvailable()) {
      throw new Error('Excel context is unavailable. Please run inside Microsoft Excel.');
    }

    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const cell = sheet.getRange(cellAddress || 'A2');
      cell.formulas = [[formula]];
      cell.select();
      await context.sync();
    });
  }

  public async insertQuoteTable(quotes: MarketQuote[], fields: string[], startCell: string = 'A1'): Promise<void> {
    if (!this.isOfficeAvailable()) {
      throw new Error('Excel context is unavailable. Please run inside Microsoft Excel.');
    }

    const fieldLabels: Record<string, string> = {
      symbol: 'Symbol',
      price: 'Price',
      change: 'Change',
      changePercent: '% Change',
      open: 'Open',
      high: 'High',
      low: 'Low',
      volume: 'Volume',
      marketCap: 'Market Cap',
      pe: 'P/E Ratio',
      week52High: '52 Week High'
    };

    const activeFields = fields.length > 0 ? fields : ['symbol', 'price', 'change', 'changePercent', 'open', 'high', 'low', 'volume', 'marketCap', 'pe'];
    const headerRow = activeFields.map(f => fieldLabels[f] || f);

    const rows = quotes.map(q => {
      return activeFields.map(f => (q as any)[f]);
    });

    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const numRows = rows.length + 1;
      const numCols = activeFields.length;

      const targetRange = sheet.getRange(startCell).getResizedRange(numRows - 1, numCols - 1);
      targetRange.values = [headerRow, ...rows];

      // Format Header
      const headerRange = sheet.getRange(startCell).getResizedRange(0, numCols - 1);
      headerRange.format.fill.color = '#005FB8'; // Accent Blue matching screenshot
      headerRange.format.font.bold = true;
      headerRange.format.font.color = '#ffffff';

      // Auto fit
      sheet.getUsedRange().format.autofitColumns();
      await context.sync();
    });
  }

  public async insertHistoryTable(history: HistoricalBar[], fields: string[], startCell: string = 'A1'): Promise<void> {
    if (!this.isOfficeAvailable()) {
      throw new Error('Excel context is unavailable.');
    }

    const fieldLabels: Record<string, string> = {
      date: 'Date',
      open: 'Open',
      high: 'High',
      low: 'Low',
      close: 'Close',
      volume: 'Volume',
      adjClose: 'Adj Close'
    };

    const activeFields = fields.length > 0 ? fields : ['date', 'open', 'high', 'low', 'close', 'volume'];
    const headerRow = activeFields.map(f => fieldLabels[f] || f);
    const rows = history.map(h => activeFields.map(f => (h as any)[f]));

    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const numRows = rows.length + 1;
      const numCols = activeFields.length;

      const targetRange = sheet.getRange(startCell).getResizedRange(numRows - 1, numCols - 1);
      targetRange.values = [headerRow, ...rows];

      const headerRange = sheet.getRange(startCell).getResizedRange(0, numCols - 1);
      headerRange.format.fill.color = '#005FB8';
      headerRange.format.font.bold = true;
      headerRange.format.font.color = '#ffffff';

      sheet.getUsedRange().format.autofitColumns();
      await context.sync();
    });
  }

  public async insertNewsTable(news: MarketNewsItem[], fields: string[], startCell: string = 'A1'): Promise<void> {
    if (!this.isOfficeAvailable()) {
      throw new Error('Excel context is unavailable.');
    }

    const fieldLabels: Record<string, string> = {
      title: 'Title',
      summary: 'Summary',
      source: 'Source',
      publishedAt: 'Published At',
      url: 'URL',
      sentiment: 'Sentiment'
    };

    const activeFields = fields.length > 0 ? fields : ['title', 'summary', 'source', 'publishedAt', 'sentiment'];
    const headerRow = activeFields.map(f => fieldLabels[f] || f);
    const rows = news.map(n => activeFields.map(f => (n as any)[f]));

    return Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const numRows = rows.length + 1;
      const numCols = activeFields.length;

      const targetRange = sheet.getRange(startCell).getResizedRange(numRows - 1, numCols - 1);
      targetRange.values = [headerRow, ...rows];

      const headerRange = sheet.getRange(startCell).getResizedRange(0, numCols - 1);
      headerRange.format.fill.color = '#005FB8';
      headerRange.format.font.bold = true;
      headerRange.format.font.color = '#ffffff';

      sheet.getUsedRange().format.autofitColumns();
      await context.sync();
    });
  }
}
