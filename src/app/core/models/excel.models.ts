export interface ProductSaleRecord {
  id: number;
  product: string;
  category: string;
  units: number;
  unitPrice: number;
  revenue: number;
  date: string;
}

export interface IExcelWorksheetRepository {
  readSelectedRange(): Promise<any[][]>;
  writeToRange(startCell: string, values: any[][]): Promise<void>;
  highlightRow(rowIndex: number): Promise<void>;
  isAvailable(): boolean;
}

export interface ISampleDataProvider {
  getSampleSalesDataset(): (string | number)[][];
}
