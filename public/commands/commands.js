/**
 * Ribbon Actions for Market Data
 */
export class MarketRibbonService {
    async refreshMarketData(event) {
        try {
            if (typeof Excel !== 'undefined') {
                await Excel.run(async (context) => {
                    context.workbook.worksheets.getActiveWorksheet().calculate(true);
                    await context.sync();
                });
            }
        }
        catch (e) {
            console.error('Refresh error:', e);
        }
        finally {
            event.completed();
        }
    }
    async insertSampleWatchlist(event) {
        try {
            if (typeof Excel !== 'undefined') {
                await Excel.run(async (context) => {
                    const sheet = context.workbook.worksheets.getActiveWorksheet();
                    const header = ['Symbol', 'Price', 'Change', '% Change', 'Open', 'High', 'Low', 'Volume', 'Market Cap', 'P/E Ratio'];
                    const rows = [
                        ['AAPL', 224.50, 2.35, 0.0106, 221.10, 225.20, 220.50, 58120345, '3.42T', 34.12],
                        ['MSFT', 415.30, -1.25, -0.0030, 416.50, 418.20, 412.10, 22345678, '3.09T', 36.45],
                        ['GOOGL', 138.21, 1.10, 0.0080, 137.10, 139.50, 136.20, 18905432, '1.74T', 27.18],
                        ['AMZN', 182.45, 0.95, 0.0052, 181.20, 184.30, 180.50, 25678901, '1.90T', 41.22]
                    ];
                    const target = sheet.getRange('A1').getResizedRange(4, 9);
                    target.values = [header, ...rows];
                    sheet.getRange('A1:J1').format.fill.color = '#005FB8';
                    sheet.getRange('A1:J1').format.font.bold = true;
                    sheet.getRange('A1:J1').format.font.color = '#ffffff';
                    sheet.getUsedRange().format.autofitColumns();
                    await context.sync();
                });
            }
        }
        catch (e) {
            console.error('Watchlist error:', e);
        }
        finally {
            event.completed();
        }
    }
}
const ribbonService = new MarketRibbonService();
function refreshMarketData(event) {
    ribbonService.refreshMarketData(event);
}
function insertSampleWatchlist(event) {
    ribbonService.insertSampleWatchlist(event);
}
if (typeof Office !== 'undefined') {
    Office.onReady(() => {
        Office.actions.associate('refreshMarketData', refreshMarketData);
        Office.actions.associate('insertSampleWatchlist', insertSampleWatchlist);
    });
}
