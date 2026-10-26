import { Injectable } from '@angular/core';
import { ISampleDataProvider } from '../models/excel.models';

/**
 * Single Responsibility: Pure provider of sample domain data
 */
@Injectable({
  providedIn: 'root'
})
export class SampleDataService implements ISampleDataProvider {
  public getSampleSalesDataset(): (string | number)[][] {
    return [
      ['ID', 'Product', 'Category', 'Units', 'Unit Price', 'Revenue', 'Date'],
      [101, 'Surface Pro 9', 'Hardware', 24, 999.99, 23999.76, '2026-03-15'],
      [102, 'Office 365 E5', 'Software', 150, 38.00, 5700.00, '2026-03-18'],
      [103, 'Azure Synapse', 'Cloud', 12, 1200.00, 14400.00, '2026-03-20'],
      [104, 'Surface Laptop Studio', 'Hardware', 18, 1599.00, 28782.00, '2026-03-22'],
      [105, 'Power BI Premium', 'Software', 45, 20.00, 900.00, '2026-03-25'],
      [106, 'Dynamics 365 CRM', 'Software', 30, 95.00, 2850.00, '2026-03-28'],
      [107, 'Azure OpenAI Service', 'Cloud', 80, 250.00, 20000.00, '2026-03-29']
    ];
  }
}
