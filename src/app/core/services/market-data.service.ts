import { Injectable } from '@angular/core';
import { IMarketDataService, MarketQuote, HistoricalBar, MarketNewsItem } from '../models/market-data.models';

/**
 * Single Responsibility: Serves real-time market data, history, and financial news
 */
@Injectable({
  providedIn: 'root'
})
export class MarketDataService implements IMarketDataService {
  private readonly quotesDb: Record<string, MarketQuote> = {
    'AAPL': {
      symbol: 'AAPL',
      price: 224.50,
      change: 2.35,
      changePercent: 1.06,
      open: 221.10,
      high: 225.20,
      low: 220.50,
      volume: 58120345,
      marketCap: '3.42T',
      pe: 34.12,
      week52High: 237.23
    },
    'MSFT': {
      symbol: 'MSFT',
      price: 415.30,
      change: -1.25,
      changePercent: -0.30,
      open: 416.50,
      high: 418.20,
      low: 412.10,
      volume: 22345678,
      marketCap: '3.09T',
      pe: 36.45,
      week52High: 468.35
    },
    'GOOGL': {
      symbol: 'GOOGL',
      price: 138.21,
      change: 1.10,
      changePercent: 0.80,
      open: 137.10,
      high: 139.50,
      low: 136.20,
      volume: 18905432,
      marketCap: '1.74T',
      pe: 27.18,
      week52High: 191.75
    },
    'AMZN': {
      symbol: 'AMZN',
      price: 182.45,
      change: 0.95,
      changePercent: 0.52,
      open: 181.20,
      high: 184.30,
      low: 180.50,
      volume: 25678901,
      marketCap: '1.90T',
      pe: 41.22,
      week52High: 201.20
    },
    'RELIANCE.NS': {
      symbol: 'RELIANCE.NS',
      price: 1342.10,
      change: 14.50,
      changePercent: 1.09,
      open: 1330.00,
      high: 1349.80,
      low: 1328.00,
      volume: 11452300,
      marketCap: '18.15T',
      pe: 24.10,
      week52High: 1608.80
    },
    'TCS.NS': {
      symbol: 'TCS.NS',
      price: 4120.50,
      change: -22.30,
      changePercent: -0.54,
      open: 4145.00,
      high: 4160.00,
      low: 4105.00,
      volume: 2314500,
      marketCap: '14.90T',
      pe: 29.80,
      week52High: 4585.00
    }
  };

  public async getQuote(symbol: string): Promise<MarketQuote> {
    const key = (symbol || '').trim().toUpperCase();
    if (this.quotesDb[key]) {
      return { ...this.quotesDb[key] };
    }
    // Synthetic fallback quote
    return {
      symbol: key,
      price: 150.00,
      change: 1.20,
      changePercent: 0.81,
      open: 149.00,
      high: 151.20,
      low: 148.50,
      volume: 5000000,
      marketCap: '1.20T',
      pe: 25.00,
      week52High: 175.00
    };
  }

  public async getQuotes(symbols: string[]): Promise<MarketQuote[]> {
    return Promise.all(symbols.map(s => this.getQuote(s)));
  }

  public async getHistory(symbol: string): Promise<HistoricalBar[]> {
    const sym = symbol.toUpperCase();
    const basePrice = this.quotesDb[sym]?.price || 200;
    const history: HistoricalBar[] = [];
    const date = new Date(2026, 9, 23); // Oct 23, 2026

    for (let i = 0; i < 15; i++) {
      const d = new Date(date);
      d.setDate(d.getDate() - i);
      const variance = (Math.sin(i) * 3) + ((i % 3) - 1);
      const close = Number((basePrice - variance).toFixed(2));
      const open = Number((close - 0.75).toFixed(2));
      const high = Number((Math.max(open, close) + 1.25).toFixed(2));
      const low = Number((Math.min(open, close) - 1.10).toFixed(2));
      const volume = Math.floor(40000000 + (Math.sin(i) * 10000000));

      history.push({
        date: d.toISOString().split('T')[0],
        open,
        high,
        low,
        close,
        volume,
        adjClose: close
      });
    }

    return history;
  }

  public async getNews(symbol: string): Promise<MarketNewsItem[]> {
    const sym = symbol.toUpperCase();
    return [
      {
        id: '1',
        title: `${sym} reports record cloud revenue growth and raises guidance`,
        summary: `Strong enterprise adoption across AI and cloud infrastructure drove outperformance for ${sym} in the latest fiscal quarter.`,
        source: 'Bloomberg Financial',
        publishedAt: '2026-10-25 09:30 AM',
        url: 'https://bloomberg.com',
        sentiment: 'positive'
      },
      {
        id: '2',
        title: `Analysts revise price targets following ${sym} product launch`,
        summary: 'Major investment banks updated their consensus estimates citing strong holiday season momentum and improved margin mix.',
        source: 'Reuters Markets',
        publishedAt: '2026-10-24 02:15 PM',
        url: 'https://reuters.com',
        sentiment: 'positive'
      },
      {
        id: '3',
        title: `Macroeconomic trends influence tech sector capital expenditures`,
        summary: 'Interest rate stability provides tailwinds for large-cap tech platforms while enterprise IT budgets expand.',
        source: 'Wall Street Journal',
        publishedAt: '2026-10-23 11:45 AM',
        url: 'https://wsj.com',
        sentiment: 'neutral'
      }
    ];
  }
}
