import { Injectable } from '@angular/core';
import { ExcelWorksheetRepository } from './excel-worksheet.repository';
import { SampleDataService } from './sample-data.service';
import { ProductSaleRecord } from '../models/excel.models';

/**
 * Dependency Inversion: OfficeFacadeService orchestrates high-level business logic
 * by delegating to specialized services.
 */
@Injectable({
  providedIn: 'root'
})
export class OfficeFacadeService {
  constructor(
    private worksheetRepo: ExcelWorksheetRepository,
    private sampleDataProvider: SampleDataService
  ) {}

  public isOfficeConnected(): boolean {
    return this.worksheetRepo.isAvailable();
  }

  public async getSelectionData(): Promise<ProductSaleRecord[]> {
    const rawValues = await this.worksheetRepo.readSelectedRange();
    if (!rawValues || rawValues.length < 2) {
      throw new Error('Selection must have at least 2 rows (header + data).');
    }

    const records: ProductSaleRecord[] = [];
    for (let i = 1; i < rawValues.length; i++) {
      const row = rawValues[i];
      if (row && row.length >= 6) {
        records.push({
          id: Number(row[0]) || (100 + i),
          product: String(row[1] || 'Product ' + i),
          category: String(row[2] || 'General'),
          units: Number(row[3]) || 0,
          unitPrice: Number(row[4]) || 0,
          revenue: Number(row[5]) || 0,
          date: String(row[6] || new Date().toISOString().split('T')[0])
        });
      }
    }
    return records;
  }

  public async exportRecordsToSheet(records: ProductSaleRecord[]): Promise<void> {
    const headers = ['ID', 'Product', 'Category', 'Units', 'Unit Price', 'Revenue', 'Date'];
    const rows = records.map((item) => [
      item.id,
      item.product,
      item.category,
      item.units,
      item.unitPrice,
      item.revenue,
      item.date
    ]);

    await this.worksheetRepo.writeToRange('A1', [headers, ...rows]);
  }

  public async injectSampleData(): Promise<void> {
    const sample = this.sampleDataProvider.getSampleSalesDataset();
    await this.worksheetRepo.writeToRange('A1', sample);
  }

  public async highlightRow(rowIndex: number): Promise<void> {
    await this.worksheetRepo.highlightRow(rowIndex);
  }
}
