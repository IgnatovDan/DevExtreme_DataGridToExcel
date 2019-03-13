/// <summary>
/// Exports the passed dxDataGrid widget into XLSX file with the help of the ExcelJS library.
/// </summary>

/// <param name="dataGrid">
/// A dxDataGrid widget.
/// Required parameter.
/// </param>

/// <param name="workbook">
/// A workbook object.
/// Optional parameter. A new ExcelJS workbook object will be created in the 'workbook' object if its value is null/undefined. The ExcelJS worksheet object will be created as well. Note that in this case it is not possible to customize the created objects and not possible to generate and save the XLSX file manually.
/// Pass a workbook object if you created it in your code and want to automatically generate and save the XLSX file.
/// See https://github.com/exceljs/exceljs#create-a-workbook for more details.
/// </param>

/// <param name="worksheet">
/// A worksheet object. Pass a workbook object if you created it in your code and want to automatically generate and save the XLSX file.
/// Optional parameter. A new ExcelJS worksheet object will be created in the 'workbook' object if its value is null/undefined. Note that in this case it is not possible to customize the created 'worksheet' object.
/// See https://github.com/exceljs/exceljs#add-a-worksheet for more details.
/// </param>

/// <param name="fileName">
/// A name of the automatically generated XLSX file.
/// Optional parameter. Default value is 'DataGrid.xlsx'.
/// See https://github.com/exceljs/exceljs#add-a-worksheet for more details.
/// </param>

/// <param name="topLeftCell">
/// A { row, column } object that specifies the top left cell starting at which the dxDataGrid content will be exported.
/// Optional parameter. Default value is { row: 1, column: 1 }.
/// Note that the both the 'row' and 'column' values are /*1-based*/ indexes.
/// </param>

/// <param name="customizeCell">
/// A callback that will be called for each generated ExcelJS cell (except of empty ones).
/// Optional parameter.
/// A { dataGrid, gridCell, cell } object will be passed to this callback.
/// The 'dataGrid' property refers to the dxDataGrid widget.
/// The 'gridCell' property is an { rowType, column, data, groupSummaryItems, totalSummaryItemName, value } object that describes the dxDataGrid cell that is being currently exported.
/// The 'cell' property refers to an ExcelJS cell that represents the 'gridCell' dxDataGrid cell.
/// </param>

/// <result>
/// A Promise object if the automatical XLSX file generating and saving is disabled. Otherwise 'undefined'.
/// A { from: { row, column }, to: { row, column } } object will be passed to the 'resolve' method of the returned promise. This object describes the range of Excel cells with dxDataGrid data.
/// Note that the both the 'row' and 'column' values are /*1-based*/ indexes.
/// </result>

function exportDataGrid({ dataGrid, workbook, worksheet, fileName = "DataGrid.xlsx", 
                         topLeftCell = { row: 1 /*1-based*/, column: 1 /*1-based*/ },
                         saveEnabled = true, customizeCell }) {
    if (!dataGrid) {
        throw "Incorrect arguments: 'dataGrid' is null";
    }
    if(!saveEnabled && (!worksheet || !workbook)) {
        throw "Incorrect arguments: 'workbook' and 'worksheet' are null when 'saveEnabled' is false";
    }
    if (!worksheet) {
        workbook = new ExcelJS.Workbook();
        worksheet = workbook.addWorksheet('Sheet 1');
    }
    if(saveEnabled && !workbook) {
        throw "Incorrect arguments: 'workbook' is null when 'saveEnabled' is true";
    }
    var result = { from: { ...topLeftCell }, to: { ...topLeftCell } };

    var currentColumnIndex = result.from.column;
    var headerRow = worksheet.getRow(result.to.row);
    for (let i = 0; i < dataGrid.getVisibleColumns().length; i++) {
        headerRow.getCell(currentColumnIndex).value = dataGrid.getVisibleColumns()[i].caption;
        customizeCell && customizeCell({
            dataGrid,
            cell: headerRow.getCell(currentColumnIndex),
            gridCell: {
                rowType: "header",
                // TODO: column, data, groupSummaryItems, totalSummaryItemName, value
            }
        });
        currentColumnIndex++;
    }
    result.to.row++;
    result.to.column += dataGrid.getVisibleColumns().length;

    return dataGrid.getController("data").loadAll().then(
        function (items) {
            for (let i = 0; i < items.length; i++) {
                var dataRow = worksheet.getRow(result.to.row);
                currentColumnIndex = result.from.column;
                for (let j = 0; j < items[i].values.length; j++) {
                    dataRow.getCell(currentColumnIndex).value = items[i].values[j];
                    customizeCell && customizeCell({
                        dataGrid,
                        cell: dataRow.getCell(currentColumnIndex),
                        gridCell: {
                            rowType: "data",
                            // TODO: column, data, groupSummaryItems, totalSummaryItemName, value
                        }
                    });
                    currentColumnIndex++;
                }
                result.to.row++;
            }
            result.to.row--;
            return Promise.resolve(result);
        }).then(function(range) {
            if(saveEnabled && workbook) {
                return workbook.xlsx.writeBuffer().then(function (buffer) {
                    var localFileName = fileName || "DataGrid.xlsx";
                    if (localFileName.substring(localFileName.length - ".xlsx".length, localFileName.length) !== ".xlsx") {
                        localFileName += ".xlsx";
                    }
                    saveAs(
                        new Blob([buffer], { type: "application/octet-stream" }),
                        localFileName
                    );
                });
            } else {
                return Promise.resolve(range);
            }
        });
}

window.exportDataGrid = exportDataGrid;
