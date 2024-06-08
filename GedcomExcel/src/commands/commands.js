Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Assign highlight function to the button
    Office.actions.associate("highlightSelectedCell", highlightSelectedCell);
  }
});

function highlightSelectedCell() {
  Excel.run(function (context) {
    // Get the active worksheet
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // Get the selected range
    const range = context.workbook.getSelectedRange();

    // Load the range address for later use
    range.load("address");

    // Highlight the selected range
    range.format.fill.color = "yellow";

    // Sync the context to apply the changes
    return context.sync().then(function () {
      console.log("Selected cell(s) highlighted: " + range.address);
    });
  }).catch(function (error) {
    console.error(error);
  });
}