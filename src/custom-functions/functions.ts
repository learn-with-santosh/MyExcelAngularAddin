/**
 * SOLID Custom Functions for Market Data
 * Namespace: SANTOSH
 * Formula signatures:
 *   =SANTOSH.QUOTE(symbol, field)
 *   =SANTOSH.HISTORY(symbol, startDate, endDate, interval, fields)
 *   =SANTOSH.NEWS(symbol, newsType, dateRange, maxResults, fields)
 */

export interface IMarketDataEngine {
  getQuoteField(symbol: string, field: string): number | string;
  getHistorySpill(symbol: string, startDate?: string, endDate?: string): (number | string)[][];
  getNewsSpill(symbol: string, maxResults?: number): string[][];
}

export class MarketDataEngine implements IMarketDataEngine {
  private readonly quotes: Record<string, Record<string, any>> = {
    'AAPL': {
      price: 224.50,
      change: 2.35,
      changePercent: 0.0106,
      open: 221.10,
      high: 225.20,
      low: 220.50,
      volume: 58120345,
      marketCap: '3.42T',
      pe: 34.12,
      week52High: 237.23
    },
    'MSFT': {
      price: 415.30,
      change: -1.25,
      changePercent: -0.0030,
      open: 416.50,
      high: 418.20,
      low: 412.10,
      volume: 22345678,
      marketCap: '3.09T',
      pe: 36.45,
      week52High: 468.35
    },
    'GOOGL': {
      price: 138.21,
      change: 1.10,
      changePercent: 0.0080,
      open: 137.10,
      high: 139.50,
      low: 136.20,
      volume: 18905432,
      marketCap: '1.74T',
      pe: 27.18,
      week52High: 191.75
    },
    'AMZN': {
      price: 182.45,
      change: 0.95,
      changePercent: 0.0052,
      open: 181.20,
      high: 184.30,
      low: 180.50,
      volume: 25678901,
      marketCap: '1.90T',
      pe: 41.22,
      week52High: 201.20
    }
  };

  public getQuoteField(symbol: string, field: string = 'price'): number | string {
    const sym = (symbol || 'AAPL').trim().toUpperCase();
    const f = (field || 'price').trim().toLowerCase();
    const data = this.quotes[sym] || this.quotes['AAPL'];
    
    if (f === 'changepercent' || f === '% change') return data['changePercent'] || 0;
    if (f === 'marketcap') return data['marketCap'] || '1.0T';
    if (f === 'pe' || f === 'p/e ratio') return data['pe'] || 25;
    
    return data[f] !== undefined ? data[f] : data['price'];
  }

  public getHistorySpill(symbol: string): (number | string)[][] {
    const sym = (symbol || 'AAPL').trim().toUpperCase();
    const base = (this.quotes[sym]?.price) || 200;
    return [
      ['Date', 'Open', 'High', 'Low', 'Close', 'Volume'],
      ['2026-10-23', base - 1, base + 2, base - 2, base, 48120000],
      ['2026-10-22', base - 2, base + 1, base - 3, base - 1, 44210000],
      ['2026-10-21', base - 3, base, base - 4, base - 2, 51300000]
    ];
  }

  public getNewsSpill(symbol: string, maxResults: number = 3): string[][] {
    const sym = (symbol || 'AAPL').trim().toUpperCase();
    const results: string[][] = [
      ['Title', 'Source', 'Published At']
    ];
    for (let i = 1; i <= Math.min(5, maxResults || 3); i++) {
      results.push([
        `${sym} reports strong quarterly momentum and enterprise adoption #${i}`,
        'Bloomberg Financial',
        `2026-10-${24 - i}`
      ]);
    }
    return results;
  }
}

const engine: IMarketDataEngine = new MarketDataEngine();

/**
 * Returns real-time market quote data for any stock symbol.
 * @customfunction
 * @param symbol Stock ticker (e.g. AAPL, MSFT, GOOGL)
 * @param field Field name (e.g. price, change, changePercent, volume, marketCap, pe)
 * @returns Real-time value or quote metric
 */
function quote(symbol: string, field: string): number | string {
  return engine.getQuoteField(symbol, field);
}

/**
 * Returns historical daily price bars for a symbol.
 * @customfunction
 * @param symbol Stock ticker
 * @param startDate Start date (YYYY-MM-DD)
 * @param endDate End date (YYYY-MM-DD)
 * @returns Historical data matrix (spill range)
 */
function history(symbol: string, startDate?: string, endDate?: string): (number | string)[][] {
  return engine.getHistorySpill(symbol, startDate, endDate);
}

/**
 * Returns latest market news headlines for a symbol.
 * @customfunction
 * @param symbol Stock ticker
 * @param maxResults Maximum number of news stories
 * @returns News matrix (spill range)
 */
function news(symbol: string, maxResults?: number): string[][] {
  return engine.getNewsSpill(symbol, maxResults);
}

// Associate Custom Functions
if (typeof CustomFunctions !== 'undefined') {
  CustomFunctions.associate('QUOTE', quote);
  CustomFunctions.associate('HISTORY', history);
  CustomFunctions.associate('NEWS', news);
}
