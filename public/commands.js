/* global Office, Excel */

Office.onReady(() => {
  // Office is ready
});

/**
 * Ribbon action: Insert quick timestamp into current active cell
 * @param event Office.AddinCommands.Event
 */
function insertTimestamp(event) {
  Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.values = [[new Date().toLocaleString()]];
    range.format.autofitColumns();
    await context.sync();
  }).then(() => {
    event.completed();
  }).catch((error) => {
    console.error('Ribbon Action Error: ', error);
    event.completed();
  });
}

/**
 * Ribbon action: Highlight currently selected cells with yellow background
 * @param event Office.AddinCommands.Event
 */
function highlightSelection(event) {
  Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.format.fill.color = '#FFF275';
    await context.sync();
  }).then(() => {
    event.completed();
  }).catch((error) => {
    console.error('Ribbon Action Error: ', error);
    event.completed();
  });
}

// Register ribbon functions in Office runtime
Office.actions.associate('insertTimestamp', insertTimestamp);
Office.actions.associate('highlightSelection', highlightSelection);
