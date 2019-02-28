"use strict";

/* eslint no-console:off */

// Load the input workbook from file.
const XlsxPopulate = require('../../lib/XlsxPopulate');

// Get template workbook and sheet.
XlsxPopulate.fromFileAsync('./template.xlsx')
    .then(workbook => {
        workbook.cloneSheet(workbook.activeSheet(), 'Cloned Sheet');

        return workbook.toFileAsync('./out.xlsx');
    })
    .catch(err => console.error(err));
