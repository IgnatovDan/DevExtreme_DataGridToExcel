function exportDataGrid({ dataGrid, worksheet, topLeftCell = { row: 1 /*1-based*/, column: 1 /*1-based*/} }) {
    if(!dataGrid) {
      throw "Incorrect arguments: 'dataGrid' is null";
    }
    var workbook;
    if(!worksheet) {
      workbook = new ExcelJS.Workbook();    
      worksheet = workbook.addWorksheet('Sheet 1');
    }
    var result = { from: { ...topLeftCell }, to: { ...topLeftCell } };
  
    var currentColumnIndex = result.from.column;
    var headerRow = worksheet.getRow(result.to.row);
    for(let i = 0; i < dataGrid.getVisibleColumns().length; i++) {
      headerRow.getCell(currentColumnIndex).value = dataGrid.getVisibleColumns()[i].caption;
      currentColumnIndex++;
    }
    result.to.row++;
    result.to.column += dataGrid.getVisibleColumns().length;
  
    return dataGrid.getController("data").loadAll().then(
      function (items) {
        for(let i = 0; i < items.length; i++) {
          var dataRow = worksheet.getRow(result.to.row);
          currentColumnIndex = result.from.column;
          for(let j = 0; j < items[i].values.length; j++) {
            dataRow.getCell(currentColumnIndex).value = items[i].values[j];
            currentColumnIndex++;
          }
          result.to.row++;
        }
        result.to.row--;
        return Promise.resolve(result);
      }).then(function() {
        if(workbook) {
          return workbook.xlsx.writeBuffer().then(function(buffer) { 
            saveAs( 
              new Blob([buffer], { type: "application/octet-stream" }), 
              "DataGrid.xlsx"
            ); 
          });
        } else {
          return Promise.resolve();
        }
      });
  }

  window.exportDataGrid = exportDataGrid;
  