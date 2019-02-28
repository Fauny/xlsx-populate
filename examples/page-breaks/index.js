"use strict";

const XlsxPopulate = require('../../lib/XlsxPopulate');

// Load the input workbook from file.
XlsxPopulate.fromBlankAsync()
    .then(workbook => {
        // Modify the workbook.
        const sheet = workbook.sheet("Sheet1");
        sheet.range('A1:B2').value([['Page1', 'Page3'], ['Page2', 'Page4']]);
        sheet.column('A').addPageBreak();
        sheet.cell('A1').addPageBreak();

        // Write to file.
        return workbook.toFileAsync("./out.xlsx");
    })
    .catch(err => console.error(err));
