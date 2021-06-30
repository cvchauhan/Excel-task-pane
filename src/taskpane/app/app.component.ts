import { Component } from "@angular/core";
// images references in the manifest
import "../../../assets/icon-16.png";
import "../../../assets/icon-32.png";
import "../../../assets/icon-80.png";
const template = require("./app.component.html");
/* global console, Excel, require */

@Component({
  selector: "app-home",
  template,
})
export default class AppComponent {
  dialog = null;
  welcomeMessage = "Welcome";
  FreezeHeader = "Freeze Header";
  CreateTable = "Create Table";  
  FilterTable = "Filter Table";  
  SortTable = "Sort Table";    
  CreateChart = "Create Chart";  
  async run() {
    try {
      await Excel.run(async (context) => {
        /**
         * Insert your Excel code here
         */
        const range = context.workbook.getSelectedRange();

        // Read the range address
        range.load("address");

        // Update the fill color
        range.format.fill.color = "red";

        await context.sync();
        console.log(`The range address was ${range.address}.`);
      });
    } catch (error) {
      console.error(error);
    }
  }
  createTable() {
      Excel.run(function (context) {        
        // TODO1: Queue table creation logic here.
        const arr = [["Date", "Merchant", "Category", "Amount"]];
        var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();                
        var expensesTable = currentWorksheet.tables.add("A1:D1", true /*hasHeaders*/);
        expensesTable.name = "ExpensesTable";
        expensesTable.getHeaderRowRange().values = arr;        
          // TODO2: Queue commands to populate the table with data.
          expensesTable.rows.add(null /*add at the end*/, [
            ["1/1/2017", "The Phone Company", "Communications", "120"],
            ["1/2/2017", "Northwind Electric Cars", "Transportation", "142.33"],
            ["1/5/2017", "Best For You Organics Company", "Groceries", "27.9"],
            ["1/10/2017", "Coho Vineyard", "Restaurant", "33"],
            ["1/11/2017", "Bellows College", "Education", "350.1"],
            ["1/15/2017", "Trey Research", "Other", "135"],
            ["1/15/2017", "Best For You Organics Company", "Groceries", "97.88"]
        ]);  
          // TODO3: Queue commands to format the table.
          expensesTable.columns.getItemAt(3).getRange().numberFormat = [['\u20AC#,##0.00']];
          expensesTable.getRange().format.autofitColumns();
          expensesTable.getRange().format.autofitRows();
          return context.sync();
      })
      .catch(function (error) {
          console.log("Error: " + error);
          if (error instanceof OfficeExtension.Error) {
              console.log("Debug info: " + JSON.stringify(error.debugInfo));
          }
      });
  }
  sortTable() {
      Excel.run(function (context) {
        var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
        var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
        var sortFields = [
            {
                key: 1,            // Merchant column
                ascending: false,
            }
        ];

        expensesTable.sort.apply(sortFields);
          // TODO1: Queue commands to sort the table by Merchant name.

          return context.sync();
      })
      .catch(function (error) {
          console.log("Error: " + error);
          if (error instanceof OfficeExtension.Error) {
              console.log("Debug info: " + JSON.stringify(error.debugInfo));
          }
      });
  }
  filterTable() {
    Excel.run(function (context) {
      var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
      var categoryFilter = expensesTable.columns.getItem('Category').filter;
      categoryFilter.applyValuesFilter(['Education', 'Groceries']);
        // TODO1: Queue commands to filter out all expense categories except
        //        Groceries and Education.

        return context.sync();
    })
    .catch(function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
  }
  createChart() {
    Excel.run(function (context) {
      var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      var expensesTable = currentWorksheet.tables.getItem('ExpensesTable');
      var dataRange = expensesTable.getDataBodyRange();
      var chart = currentWorksheet.charts.add('ColumnClustered', dataRange, 'Auto');
      chart.setPosition("A15", "F30");
      chart.title.text = "Expenses";
      chart.legend.position = "Right";
      chart.legend.format.fill.setSolidColor("red");
      chart.dataLabels.format.font.size = 15;
      chart.dataLabels.format.font.color = "black";
      chart.series.getItemAt(0).name = 'Value in \u20AC';
        // TODO1: Queue commands to get the range of data to be charted.

        // TODO2: Queue command to create the chart and define its type.

        // TODO3: Queue commands to position and format the chart.

        return context.sync();
    })
    .catch(function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
  }
  freezeHeader() {
    Excel.run(function (context) {
      var currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
      currentWorksheet.freezePanes.freezeRows(1);
        // TODO1: Queue commands to keep the header visible when the user scrolls.

        return context.sync();
    })
    .catch(function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
  }
  openDialog() {    
    // TODO1: Call the Office Common API that opens a dialog
    Office.context.ui.displayDialogAsync(
      'https://localhost:3000/popup.html',
      {height: 45, width: 55},
  
      function (result) {
        this.dialog = result.value;
        this.dialog.addEventHandler(Office.EventType.DialogMessageReceived, this.processMessage);
      }      
    );
  }
  processMessage(arg) {
    document.getElementById("user-name").innerHTML = arg.message;
    this.dialog.close();
  }
}
