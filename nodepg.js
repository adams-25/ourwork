const { ipcMain } = require('electron');
const ExcelJS = require('exceljs');
const path = require('path');

function initializeIpcHandlers(mainWindow) {  // Accept mainWindow as a parameter
  ipcMain.handle('open-excel', async () => {
    try {
      const workbook = new ExcelJS.Workbook();
      const filePath = 'D:\\x\\db.xlsx';  // Path to your Excel file
      await workbook.xlsx.readFile(filePath);

      const worksheet = workbook.getWorksheet('ProH');
      console.log(worksheet.name);

      // Loop through rows starting from row 2 to the last row with data
      let projectData = [];
      worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        if (rowNumber > 1) {  // Skip the first row (header)
          const col1 = row.getCell(1).value;  // Get value from column 1
          const col2 = row.getCell(2).value;  // Get value from column 2
          projectData.push(`${col1} - ${col2}`);
        }
      });

      // Send the data to the renderer to update the treeview
      mainWindow.webContents.send('update-treeview', projectData);

      return 'Done ...............................';
    } catch (error) {
      console.error('Error:', error);
      return 'Error occurred while opening the Excel file.';
    }
  });
}

module.exports = {
  initializeIpcHandlers
};
