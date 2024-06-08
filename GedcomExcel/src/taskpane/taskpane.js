Office.onReady(function() {
    // Office is ready.
    $(document).ready(function () {
        // The document is ready.
    });
});

function sayHello() {
    Excel.run((context) => {
      context.workbook.worksheets.getActiveWorksheet().getRange('A1').values = [
        ['Hello world!'],
      ];
      return context.sync();
    });
  }