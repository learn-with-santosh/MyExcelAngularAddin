import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';

// Services & Models
import { MarketDataService } from './core/services/market-data.service';
import { ExcelMarketWriter } from './core/services/excel-market-writer.service';
import { MarketQuote, HistoricalBar, MarketNewsItem, ReadyTemplate } from './core/models/market-data.models';

// Kendo UI Angular Modules
import { GridModule } from '@progress/kendo-angular-grid';
import { ButtonModule } from '@progress/kendo-angular-buttons';
import { InputsModule } from '@progress/kendo-angular-inputs';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [
    CommonModule,
    FormsModule,
    GridModule,
    ButtonModule,
    InputsModule
  ],
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent implements OnInit {
  // Navigation tabs matching screenshot: Quote | History | News | Templates
  public activeTab: 'quote' | 'history' | 'news' | 'templates' = 'quote';
  public isConnectedToExcel = false;
  public statusToast: string | null = null;
  public showPreviewModal = false;
  public previewData: any[] = [];
  public previewTitle = 'Data Preview';

  // --- TAB 1: QUOTE STATE ---
  public quoteSymbol: string = 'AAPL';
  public quoteOutputOption: 'single' | 'table' = 'single';
  public insertLocation: string = 'A2';
  public quoteFields: Record<string, boolean> = {
    price: true,
    change: true,
    changePercent: true,
    open: false,
    high: false,
    low: false,
    volume: true,
    marketCap: false,
    pe: false,
    week52High: false
  };

  // --- TAB 2: HISTORY STATE ---
  public historySymbol: string = 'AAPL';
  public historyDateRange: string = 'Custom Range';
  public historyStartDate: string = '2026-10-01';
  public historyEndDate: string = '2026-10-23';
  public historyInterval: string = '1 Day';
  public historyOutputOption: 'table' | 'formula' = 'table';
  public historyFields: Record<string, boolean> = {
    date: true,
    open: true,
    high: true,
    low: true,
    close: true,
    volume: true,
    adjClose: false
  };

  // --- TAB 3: NEWS STATE ---
  public newsSymbol: string = 'AAPL';
  public newsType: string = 'All News';
  public newsDateRange: string = 'Last 7 Days';
  public newsMaxResults: number = 10;
  public newsOutputOption: 'table' | 'formula' = 'table';
  public newsFields: Record<string, boolean> = {
    title: true,
    summary: true,
    source: true,
    publishedAt: true,
    url: true,
    sentiment: true
  };

  // --- TAB 4: TEMPLATES ---
  public templates: ReadyTemplate[] = [
    {
      id: 'quote-sheet',
      title: 'Quick Quote Sheet',
      description: 'Create a ready-to-use quote table with common fields.',
      icon: 'table'
    },
    {
      id: 'watchlist',
      title: 'Multi-Symbol Watchlist',
      description: 'Track multiple symbols with real-time data.',
      icon: 'list'
    },
    {
      id: 'chart-data',
      title: 'Historical Chart Data',
      description: 'Get historical data with a chart-ready format.',
      icon: 'chart'
    },
    {
      id: 'news-feed',
      title: 'Latest News Feed',
      description: 'Insert latest news for a symbol.',
      icon: 'news'
    },
    {
      id: 'overview',
      title: 'Market Overview',
      description: 'Key indices, market movers and overview.',
      icon: 'dashboard'
    }
  ];

  constructor(
    private marketData: MarketDataService,
    private excelWriter: ExcelMarketWriter
  ) {}

  public async ngOnInit(): Promise<void> {
    this.isConnectedToExcel = this.excelWriter.isOfficeAvailable();
    if (this.isConnectedToExcel) {
      try {
        const address = await this.excelWriter.getSelectedCellAddress();
        if (address) this.insertLocation = address;
      } catch {
        this.insertLocation = 'A2';
      }
    }
  }

  // --- Dynamic Formula Previews (matches screenshot) ---
  public get quoteFormulaPreview(): string {
    const sym = (this.quoteSymbol || 'AAPL').toUpperCase();
    const selected = Object.keys(this.quoteFields).filter(k => this.quoteFields[k]);
    if (this.quoteOutputOption === 'single') {
      const firstField = selected[0] || 'price';
      return `=SANTOSH.QUOTE("${sym}", "${firstField}")`;
    }
    const fieldsFormatted = selected.map(f => `"${f}"`).join(', ');
    return `=SANTOSH.QUOTE("${sym}", {${fieldsFormatted}})`;
  }

  public get historyFormulaPreview(): string {
    const sym = (this.historySymbol || 'AAPL').toUpperCase();
    const selected = Object.keys(this.historyFields).filter(k => this.historyFields[k]);
    const fieldsFormatted = selected.map(f => `"${f}"`).join(', ');
    return `=SANTOSH.HISTORY("${sym}", "${this.historyStartDate}", "${this.historyEndDate}", "1d", {${fieldsFormatted}})`;
  }

  public get newsFormulaPreview(): string {
    const sym = (this.newsSymbol || 'AAPL').toUpperCase();
    const selected = Object.keys(this.newsFields).filter(k => this.newsFields[k]);
    const fieldsFormatted = selected.map(f => `"${f}"`).join(', ');
    return `=SANTOSH.NEWS("${sym}", "all", "7d", ${this.newsMaxResults}, {${fieldsFormatted}})`;
  }

  // --- Selection Helpers ---
  public selectAllQuoteFields(): void {
    const allSelected = Object.values(this.quoteFields).every(v => v);
    Object.keys(this.quoteFields).forEach(k => {
      this.quoteFields[k] = !allSelected;
    });
  }

  // --- Insertion Actions ---
  public async executeQuoteAction(): Promise<void> {
    try {
      if (this.quoteOutputOption === 'single') {
        const formula = this.quoteFormulaPreview;
        await this.excelWriter.insertFormula(this.insertLocation, formula);
        this.showToast(`Inserted formula ${formula} at ${this.insertLocation}`);
      } else {
        const sym = (this.quoteSymbol || 'AAPL').toUpperCase();
        const quote = await this.marketData.getQuote(sym);
        const selected = Object.keys(this.quoteFields).filter(k => this.quoteFields[k]);
        await this.excelWriter.insertQuoteTable([quote], selected, this.insertLocation || 'A1');
        this.showToast(`Inserted quote table for ${sym} at ${this.insertLocation}`);
      }
    } catch (err: any) {
      this.showToast(err.message || 'Error inserting into Excel');
    }
  }

  public async executeHistoryAction(): Promise<void> {
    try {
      if (this.historyOutputOption === 'formula') {
        const formula = this.historyFormulaPreview;
        await this.excelWriter.insertFormula(this.insertLocation || 'A1', formula);
        this.showToast(`Inserted formula ${formula}`);
      } else {
        const history = await this.marketData.getHistory(this.historySymbol);
        const selected = Object.keys(this.historyFields).filter(k => this.historyFields[k]);
        await this.excelWriter.insertHistoryTable(history, selected, this.insertLocation || 'A1');
        this.showToast(`Inserted historical dataset for ${this.historySymbol}`);
      }
    } catch (err: any) {
      this.showToast(err.message || 'Error inserting history');
    }
  }

  public async executeNewsAction(): Promise<void> {
    try {
      if (this.newsOutputOption === 'formula') {
        const formula = this.newsFormulaPreview;
        await this.excelWriter.insertFormula(this.insertLocation || 'A1', formula);
        this.showToast(`Inserted news formula ${formula}`);
      } else {
        const news = await this.marketData.getNews(this.newsSymbol);
        const selected = Object.keys(this.newsFields).filter(k => this.newsFields[k]);
        await this.excelWriter.insertNewsTable(news, selected, this.insertLocation || 'A1');
        this.showToast(`Inserted latest news for ${this.newsSymbol}`);
      }
    } catch (err: any) {
      this.showToast(err.message || 'Error inserting news');
    }
  }

  public async insertTemplate(templateId: string): Promise<void> {
    try {
      if (templateId === 'watchlist' || templateId === 'quote-sheet') {
        const quotes = await this.marketData.getQuotes(['AAPL', 'MSFT', 'GOOGL', 'AMZN']);
        await this.excelWriter.insertQuoteTable(quotes, ['symbol', 'price', 'change', 'changePercent', 'open', 'high', 'low', 'volume', 'marketCap', 'pe'], 'A1');
        this.showToast('Watchlist template inserted at A1');
      } else if (templateId === 'chart-data') {
        const history = await this.marketData.getHistory('AAPL');
        await this.excelWriter.insertHistoryTable(history, ['date', 'open', 'high', 'low', 'close', 'volume'], 'A1');
        this.showToast('Chart Data template inserted at A1');
      } else if (templateId === 'news-feed') {
        const news = await this.marketData.getNews('AAPL');
        await this.excelWriter.insertNewsTable(news, ['title', 'summary', 'source', 'publishedAt'], 'A1');
        this.showToast('News Feed template inserted at A1');
      } else {
        const quotes = await this.marketData.getQuotes(['AAPL', 'MSFT', 'RELIANCE.NS', 'TCS.NS']);
        await this.excelWriter.insertQuoteTable(quotes, ['symbol', 'price', 'change', 'changePercent', 'volume'], 'A1');
        this.showToast('Market Overview template inserted at A1');
      }
    } catch (err: any) {
      this.showToast(err.message || 'Error inserting template');
    }
  }

  // --- Preview Actions ---
  public async previewQuoteData(): Promise<void> {
    const sym = (this.quoteSymbol || 'AAPL').toUpperCase();
    const quote = await this.marketData.getQuote(sym);
    this.previewTitle = `Market Quote Preview: ${sym}`;
    this.previewData = [quote];
    this.showPreviewModal = true;
  }

  public async previewHistoryData(): Promise<void> {
    const sym = (this.historySymbol || 'AAPL').toUpperCase();
    this.previewData = await this.marketData.getHistory(sym);
    this.previewTitle = `Historical Data Preview: ${sym}`;
    this.showPreviewModal = true;
  }

  public async previewNewsData(): Promise<void> {
    const sym = (this.newsSymbol || 'AAPL').toUpperCase();
    this.previewData = await this.marketData.getNews(sym);
    this.previewTitle = `Latest News Preview: ${sym}`;
    this.showPreviewModal = true;
  }

  public copyToClipboard(text: string): void {
    navigator.clipboard?.writeText(text);
    this.showToast('Formula copied to clipboard!');
  }

  private showToast(msg: string): void {
    this.statusToast = msg;
    setTimeout(() => {
      this.statusToast = null;
    }, 3500);
  }
}
