const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const ExcelJS = require('exceljs');
const fs = require('fs');

const nodepg = require('./nodepg.js');



let mainWindow;
let workbook;
let worksheet;
let directoryPath = '';
let treeViewData = [];







app.whenReady().then(() => {
    // Initialize the main window
    mainWindow = new BrowserWindow({
        webPreferences: {
            preload: path.join(__dirname, 'preload.js'),
            nodeIntegration: true,
            contextIsolation: false
   
        }
        
    });




    // Load the HTML page
    mainWindow.loadFile('index.html');

 


    
    nodepg.initializeIpcHandlers(); 
    
    app.on('window-all-closed', () => {
        if (process.platform !== 'darwin') {
          app.quit();
        }
      });


    // Initialize the database and then initialize the TreeView
    initializeDatabase().then(() => {
        initializeTreeView();
    });
});

// Function to initialize the database connection
async function initializeDatabase() {
    const filePath = 'D:\\x\\db.xlsx'

    try {
        workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(filePath);
        worksheet = workbook.getWorksheet('ProH');
        console.log('Database connected successfully');
    } catch (error) {
        console.error('Error connecting to the database:', error);
    }
}

// Function to initialize the TreeView
async function initializeTreeView() {
    if (!worksheet) {
        console.error('Worksheet is not loaded');
        return;
    }

    // Add the 'Projects' head node
    let projectsNode = { label: 'Projects', children: [] };

    // Loop through the rows and add sub-nodes to the 'Projects' node
    for (let rowIndex = 3; rowIndex <= worksheet.rowCount; rowIndex++) {
        const row = worksheet.getRow(rowIndex);
        const projectLabel = `${row.getCell(1).value} - ${row.getCell(2).value}`;

        // Add sub-node under 'Projects'
        projectsNode.children.push({ label: projectLabel });
    }

    // Send the tree data to the renderer process to update the TreeView
    console.log('Sending tree data to renderer');
    mainWindow.webContents.send('populate-treeview', [projectsNode]);
}



// IPC handlers to trigger the above functions from inv.html
ipcMain.handle('initialize-database', async () => {
    await initializeDatabase();
  });
  
  ipcMain.handle('initialize-treeview', async () => {
    await initializeTreeView();
  });

ipcMain.on('treeview-node-clicked', async (event, nodeLabel) => { 
    const firstThreeLetters = nodeLabel.substring(0, 3).toUpperCase();  // Get the first 3 letters
    console.log(`First 3 letters: ${firstThreeLetters}`);

    try {
        let rowNumber = -1;

        // Loop through the rows starting from row 3 (index 3 corresponds to row 3 in Excel)
        for (let rowIndex = 3; rowIndex <= worksheet.rowCount; rowIndex++) {
            const row = worksheet.getRow(rowIndex);
            const cellValue = row.getCell(1).value; // Get value from column A

            // Check if the value in Column A matches the first 3 letters
            if (cellValue && cellValue.toString().substring(0, 3).toUpperCase() === firstThreeLetters) {
                rowNumber = rowIndex;  // Set the row number
                break;
            }
        }

        if (rowNumber === -1) {
            console.log('No matching row found.');
            return;
        }

        console.log(`Found matching row at: ${rowNumber}`);

        // Get row data for sending to renderer
        const row = worksheet.getRow(rowNumber);
        const rowData = [];
        for (let i = 1; i <= row.cellCount; i++) {
            rowData.push(row.getCell(i).value);
        }

      
        
        // Add the rowNumber value at the end of the rowData array
        rowData.push(rowNumber);

     






        // Log the row data being sent to renderer
        console.log('Sending data to renderer:', rowData);
        console.log(rowData[37]);

        directoryPath = rowData[37]; 

        // Send the row data back to the renderer to populate the textboxes
        mainWindow.webContents.send('populate-textboxes', rowData);
       


    
     
    } catch (error) {
        console.error('Error while handling node click:', error);
    }

   














   listFilesInDirectory(directoryPath);
    // Call your listFilesInDirectory function with the updated path
   // listFilesInDirectory('G:\\Other computers\\My Laptop\\NEW ALHAZAM\\PROJECTS\\2024\\293 Federal Youth Authority - Sharjah');


    


        // Fetch data from the "srpro" sheet using the first three letters
        const srproWorksheet = workbook.getWorksheet('srpro');
        const srproData = [];
       
        // Loop through "srpro" sheet to find matches in column B
        for (let rowIndex = 3; rowIndex <= srproWorksheet.rowCount; rowIndex++) {
            const row = srproWorksheet.getRow(rowIndex);
            const cellValue = row.getCell(2).value;  // Column B in srpro sheet

            // If the value in Column B matches the first 3 letters
            if (cellValue && cellValue.toString().substring(0, 3).toUpperCase() === firstThreeLetters) {
                const rowData = [];
                for (let i = 1; i <= 8; i++) {  // From Column A to Column H
                    rowData.push(row.getCell(i).value);
                }
                srproData.push(rowData);
             
            }
        }


        //mainWindow.webContents.send('sort-and-sum', srproData);//

        // Send the "srpro" data to the renderer to populate the table
        mainWindow.webContents.send('populate-listviewtbl', srproData);

        // Send the data back to renderer for sorting and displaying the sums
        
       


});



function listFilesInDirectory(directoryPath) {
    mainWindow.webContents.send('cls-listview');

    fs.readdir(directoryPath, (err, files) => {
        if (err) {
            
            console.error('Error reading directory:', err);
           
            return;
        }

        // Send the list of files to the renderer process
        mainWindow.webContents.send('populate-listview', files);
    });
}


ipcMain.on('update-directory-path', (event, newPath) => {
    if (newPath) {
        directoryPath = newPath; // Update the directory path
        console.log('Updated directory path:', directoryPath);
    } else {
        console.error('Received invalid directory path');
    }
});

// Handle listview item click
ipcMain.on('listview-item-clicked', (event, clickedItem) => {
    console.log(`Item clicked: ${clickedItem}`);

    if (!directoryPath) {
        console.error('Directory path is not set');
        return;
    }

    // Check if the clicked item ends with 'pdf'
    if (clickedItem.slice(-3).toLowerCase() === 'pdf') {
        // Construct the full path to the PDF file using the updated directoryPath
        const filePath = `${directoryPath}\\${clickedItem}`;
        console.log(`PDF file path: ${filePath}`);

        // Send the file path to the renderer process to display it in the PDF viewer
        mainWindow.webContents.send('show-pdf', filePath);
    } else {
        console.log("Clicked item is not a PDF file");
    }
});








// Helper function to invoke the 'get-srpro-data' handler
async function getSrproData() {
    try {
       
        const sheet = workbook.Sheets['srpro'];
        return xlsx.utils.sheet_to_json(sheet, { header: 1 });
    } catch (error) {
        console.error('Error fetching srpro data:', error);
        return [];
    }
}




















// Handle the row edit request
ipcMain.handle('edit-row', async (event, lblRowCaption, inputValues) => {
  const filePath = 'D:\\x\\db.xlsx'



    try {
    
        const row = worksheet.getRow(lblRowCaption);
console.log(lblRowCaption)
        // Loop through inputValues and set cell values
        for (const { dataTag, value } of inputValues) {
            if (dataTag) {
                row.getCell(dataTag).value = value;
            }
        }


   // Commit row changes (update cell values in memory)
   row.commit();

   // Save the workbook to file
   
   await workbook.xlsx.writeFile(filePath);

   return { success: true, message: 'Row updated successfully!' };
} catch (error) {
   console.error('Error updating row:', error);
   return { success: false, message: `Failed to update row: ${error.message}` };
}
});







ipcMain.on('open-inv-page', () => {
    mainWindow.loadFile('inv.html');
});

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit();
});