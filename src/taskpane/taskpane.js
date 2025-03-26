/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office, OfficeExtension, alert, Microsoft, window */
window.addEventListener("DOMContentLoaded", () => {
  Office.onReady((info) => {
    console.log("Office is ready", info);
    if (info.host === Office.HostType.Excel) {
      const appBody = document.getElementById("app-body");
      const sideloadMsg = document.getElementById("sideload-msg");

      if (appBody) appBody.style.display = "flex";
      if (sideloadMsg) sideloadMsg.style.display = "none";

      run(); // your function to load log data
    }
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.7")) {
      console.log("Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.");
    }

    // Assign event handlers and other initialization logic.
    document.getElementById("create-table").onclick = createTable;
    // document.getElementById("filter-table").onclick = filterTable;
    // document.getElementById("sort-table").onclick = sortTable;
    document.getElementById("create-chart").onclick = createChart;
    document.getElementById("freeze-header").onclick = freezeHeader;
    document.getElementById("open-dialog").onclick = openDialog;
  });
});

// export async function run() {
//   try {
//     await Excel.run(async (context) => {
//       /**
//        * Insert your Excel code here
//        */
//       const range = context.workbook.getSelectedRange();

//       // Read the range address
//       range.load("address");

//       // Update the fill color
//       range.format.fill.color = "yellow";

//       await context.sync();
//       console.log(`The range address was ${range.address}.`);
//     });
//   } catch (error) {
//     console.error(error);
//   }
// }

// Office.onReady(() => {
//   Excel.run(async (context) => {
//     const sheet = context.workbook.worksheets.getItem("Log");
//     const range = sheet.getUsedRange();
//     range.load("values");
//     await context.sync();

//     const headers = range.values[0];
//     const rows = range.values.slice(1);
//     logEntries = rows.map((row) => {
//       const entry = {};
//       headers.forEach((h, i) => (entry[h.trim()] = row[i]));
//       return entry;
//     });

//     Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, showSelectedCellMetadata);
//     showSelectedCellMetadata(); // show initial selection
//   });
// });
function getAbsoluteCellAddress(range) {
  const row = range.rowIndex + 1; // 0-based to 1-based
  const col = String.fromCharCode("A".charCodeAt(0) + range.columnIndex); // Assume up to Z
  return `$${col}$${row}`;
}

let logEntries = [];
async function run() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Log");
    const range = sheet.getUsedRange();
    range.load("values");
    await context.sync();

    const headers = range.values[0];
    const rows = range.values.slice(1);
    logEntries = rows.map((row) => {
      const entry = {};
      headers.forEach((h, i) => (entry[h.trim()] = row[i]));
      return entry;
    });

    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, showSelectedCellMetadata);
    showSelectedCellMetadata();
  });
}
async function showSelectedCellMetadata() {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load(["rowIndex", "columnIndex", "address", "worksheet/name", "values"]);
    await context.sync();

    const sheet = range.worksheet.name;
    const cell = getAbsoluteCellAddress(range);
    const value = range.values?.[0]?.[0] ?? "(empty)";

    // Extract row and column info
    const col = cell.replace(/[0-9]/g, "");
    const row = cell.replace(/[A-Z]/gi, "");

    document.getElementById("rowNum").textContent = row;
    document.getElementById("colNum").textContent = col;
    document.getElementById("currValue").textContent = value;

    // Match log entry (optional fuzzy match support later)
    const absAddress = getAbsoluteCellAddress(range);
    const match = logEntries.find((entry) => entry["Cell"] === absAddress && entry["Sheet"] === sheet);

    if (match) {
      const note = match["Notes"] ? `<p><b>Note:</b> ${match["Notes"]}</p>` : "";
      const oldVal = match["Previous Value"] ?? "(none)";
      const newVal = match["New Value"] ?? "(none)";
      const user = match["User"] ?? "Unknown";
      document.getElementById("log-link-container").innerHTML = `
        <div class="history-item" style="background-color: #eef3ff; padding: 10px; border-radius: 8px;">
          <p><b>Previous:</b> ${oldVal}</p>
          <p><b>New:</b> ${newVal}</p>
          <p><b>User:</b> ${user}</p>
          ${note}
          <button class="ms-Button" onclick="goToLogEntry('${absAddress}', '${sheet}')">
            Go to Log Entry
          </button>
        </div>
      `;
    } else {
      document.getElementById("log-link-container").innerHTML = "<i>No log entry found for this cell.</i>";
    }
  });
}

async function goToLogEntry(cell, sheetName) {
  await Excel.run(async (context) => {
    const logSheet = context.workbook.worksheets.getItem("Log");
    const range = logSheet.getUsedRange();
    range.load("values, address");
    await context.sync();

    const headers = range.values[0];
    const rows = range.values.slice(1);

    const cellIndex = headers.indexOf("Cell");
    const sheetIndex = headers.indexOf("Sheet");

    if (cellIndex === -1 || sheetIndex === -1) {
      console.error("Log sheet is missing 'Cell' or 'Sheet' headers.");
      return;
    }

    const matchIndex = rows.findIndex((row) => row[cellIndex] === cell && row[sheetIndex] === sheetName);
    if (matchIndex !== -1) {
      logSheet.getUsedRange().format.fill.clear();

      const fullRow = logSheet.getRangeByIndexes(matchIndex + 1, 0, 1, headers.length);
      fullRow.format.fill.color = "#FFF2CC"; // Light yellow
      fullRow.select();

      await context.sync();
    } else {
      console.warn("Log entry not found.");
    }
  });
}
window.goToLogEntry = goToLogEntry;
async function freezeHeader() {
  await Excel.run(async (context) => {
    // TODO1: Queue commands to keep the header visible when the user scrolls.
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    currentWorksheet.freezePanes.freezeRows(1);

    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

async function createChart() {
  await Excel.run(async (context) => {
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.getItem("ExpensesTable");
    const dataRange = expensesTable.getDataBodyRange();

    // TODO2: Queue command to create the chart and define its type.
    const chart = currentWorksheet.charts.add("ColumnClustered", dataRange, "auto");

    // TODO3: Queue commands to position and format the chart.
    chart.setPosition("A15", "F30");
    chart.title.text = "Expenses";
    chart.legend.position = "right";
    chart.legend.format.fill.setSolidColor("white");
    chart.dataLabels.format.font.size = 15;
    chart.dataLabels.format.font.color = "black";
    chart.series.getItemAt(0).name = "Value in &euro;";

    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

async function createTable() {
  await Excel.run(async (context) => {
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    const expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
    expensesTable.name = "ExpensesTable";

    expensesTable.getHeaderRowRange().values = [["Date", "Merchant", "Category", "Amount"]];

    expensesTable.rows.add(null /*add at the end*/, [
      ["1/1/2017", "The Phone Company", "Communications", "120"],
      ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"],
      ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9"],
      ["1/10/2017", "Coho Vineyard", "Restaurant", "33"],
      ["1/11/2017", "Bellows College", "Education", "350.1"],
      ["1/15/2017", "Trey Research", "Other", "135"],
      ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88"],
    ]);
    // TODO3: Queue commands to format the table.
    expensesTable.columns.getItemAt(3).getRange().numberFormat = [["\u20AC#,##0.00"]];
    expensesTable.getRange().format.autofitColumns();
    expensesTable.getRange().format.autofitRows();
    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

// async function filterTable() {
//   await Excel.run(async (context) => {
//     // TODO1: Queue commands to filter out all expense categories except
//     //        Groceries and Education.

//     const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
//     const expensesTable = currentWorksheet.tables.getItem("ExpensesTable");
//     const categoryFilter = expensesTable.columns.getItem("Category").filter;
//     categoryFilter.applyValuesFilter(["Education", "Groceries"]);

//     await context.sync();
//   }).catch(function (error) {
//     console.log("Error: " + error);
//     if (error instanceof OfficeExtension.Error) {
//       console.log("Debug info: " + JSON.stringify(error.debugInfo));
//     }
//   });
// }

// async function sortTable() {
//   await Excel.run(async (context) => {
//     // TODO1: Queue commands to sort the table by Merchant name.
//     const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
//     const expensesTable = currentWorksheet.tables.getItem("ExpensesTable");
//     const sortFields = [
//       {
//         key: 1, // Merchant column
//         ascending: false,
//       },
//     ];

//     expensesTable.sort.apply(sortFields);
//     await context.sync();
//   }).catch(function (error) {
//     console.log("Error: " + error);
//     if (error instanceof OfficeExtension.Error) {
//       console.log("Debug info: " + JSON.stringify(error.debugInfo));
//     }
//   });
// }

async function toggleProtection(args) {
  await Excel.run(async (context) => {
    // TODO1: Queue commands to reverse the protection status of the current worksheet.
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("protection/protected");
    await context.sync();

    if (sheet.protection.protected) {
      sheet.protection.unprotect();
    } else {
      sheet.protection.protect();
    }
    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
  args.completed();
}

Office.actions.associate("toggleProtection", toggleProtection);

let dialog = null;

function openDialog() {
  // TODO1: Call the Office Common API that opens a dialog
  Office.context.ui.displayDialogAsync(
    "https://localhost:3000/popup.html",
    { height: 45, width: 55 },

    // TODO2: Add callback parameter.
    function (result) {
      dialog = result.value;
      dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, processMessage);
    }
  );
}

function processMessage(arg) {
  document.getElementById("user-name").innerHTML = arg.message;
  dialog.close();
}

// Make goToLogEntry globally accessible for button onclick
window.goToLogEntry = goToLogEntry;

// Register globals for linter
/* global alert, Microsoft, windows */
