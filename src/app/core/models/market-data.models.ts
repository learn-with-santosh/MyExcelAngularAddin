export interface MarketQuote {
  symbol: string;
  price: number;
  change: number;
  changePercent: number;
  open: number;
  high: number;
  low: number;
  volume: number;
  marketCap: string;
  pe: number;
  week52High: number;
}

export interface HistoricalBar {
  date: string;
  open: number;
  high: number;
  low: number;
  close: number;
  volume: number;
  adjClose?: number;
}

export interface MarketNewsItem {
  id: string;
  title: string;
  summary: string;
  source: string;
  publishedAt: string;
  url: string;
  sentiment: 'positive' | 'negative' | 'neutral';
}

export interface ReadyTemplate {
  id: string;
  title: string;
  description: string;
  icon: string;
}

export interface IMarketDataService {
  getQuote(symbol: string): Promise<MarketQuote>;
  getQuotes(symbols: string[]): Promise<MarketQuote[]>;
  getHistory(symbol: string, startDate?: string, endDate?: string, interval?: string): Promise<HistoricalBar[]>;
  getNews(symbol: string, newsType?: string, maxResults?: number): Promise<MarketNewsItem[]>;
}

export interface IExcelMarketWriter {
  insertQuoteTable(quotes: MarketQuote[], selectedFields: string[], startCell?: string): Promise<void>;
  insertHistoryTable(history: HistoricalBar[], selectedFields: string[], startCell?: string): Promise<void>;
  insertNewsTable(news: MarketNewsItem[], selectedFields: string[], startCell?: string): Promise<void>;
  insertFormula(cellAddress: string, formula: string): Promise<void>;
  getSelectedCellAddress(): Promise<string>;
}
