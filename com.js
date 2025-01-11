const { ipcMain } = require('electron');

const ExcelJS = require('exceljs');
const fs = require('fs');

let workbook;
let worksheet;
let srpnWorksheet;




// Function to handle node clicks
function nodeClicked() {
  const clickedNode = event.target; // Get the clicked node
  const nodeCaption = clickedNode.innerText.trim();

  // Extract the number before the "-" sign
  const number = nodeCaption.split('-')[0].trim();

  // Now perform the search logic
  findMatchingRowInSrpnSheet(number);
}

// Function to find the matching row in sheet 'srpn'
async function findMatchingRowInSrpnSheet(number) {
   srpnWorksheet = workbook.getWorksheet('srpn');
  let rowdatano = null;

  // Loop through rows of the 'srpn' sheet (starting from row 3)
  srpnWorksheet.eachRow((row, rowNumber) => {
    if (row.getCell(1).value == number) { // Match based on Column 1 value
      rowdatano = rowNumber;
    }
  });

  // If row is found, handle it
  if (rowdatano !== null) {
    const matchedRow = srpnWorksheet.getRow(rowdatano);
    
    // Collect the matched row data from columns A to H
    const matchedData = [];
    for (let i = 1; i <= 8; i++) {
      matchedData.push(matchedRow.getCell(i).value);
    }

    // Display matched data in ListView
    displayMatchedDataInListView(matchedData);
  }
}

// Function to display matched data in ListView (with id 'acctbl')
function displayMatchedDataInListView(matchedData) {
  const listView = document.getElementById('acctbl');
  const row = listView.insertRow(); // Add new row to the ListView

  // Add matched data as cells in the row
  matchedData.forEach(data => {
    const cell = row.insertCell();
    cell.innerText = data;
  });

  // Re-sort the ListView if needed (sorting by Column 1 and Column 2)
  sortListView();

  // Sum and update inputs after adding new data
  sumColumnsAndUpdateInputs();
}

// Function to sort the data in the ListView
function sortListView() {
  const rows = document.querySelectorAll('#acctbl tr');
  
  // Sort by Col1 (small to large) and then by Col2 (old date to new date)
  const sortedRows = Array.from(rows).slice(1).sort((a, b) => {
    const col1A = parseFloat(a.cells[0].innerText) || 0;
    const col1B = parseFloat(b.cells[0].innerText) || 0;
    const col2A = new Date(a.cells[1].innerText);
    const col2B = new Date(b.cells[1].innerText);

    // First sort by Col1 (small to large)
    if (col1A !== col1B) return col1A - col1B;

    // Then sort by Col2 (old date to new date)
    return col2A - col2B;
  });

  // Update the table with sorted rows
  const tbody = document.querySelector('#acctbl tbody');
  tbody.innerHTML = ''; // Clear the existing rows
  sortedRows.forEach(row => tbody.appendChild(row));
}

// Function to sum columns D, E, and H, and update input fields
function sumColumnsAndUpdateInputs() {
  let sumColD = 0;
  let sumColE = 0;
  let sumColH = 0;

  // Loop through all rows except the header (starting from row 2)
  const rows = document.querySelectorAll('#acctbl tr');
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];

    const colDValue = parseFloat(row.cells[3].innerText) || 0;
    const colEValue = new Date(row.cells[4].innerText);
    const colHValue = parseFloat(row.cells[7].innerText) || 0;

    // Sum the values of each column
    sumColD += colDValue;
    sumColE += colEValue.getTime(); // Sum as timestamps (milliseconds)
    sumColH += colHValue;
  }

  // Format the summed data for Column E (date format)
  const sumColEDate = new Date(sumColE);
  
  // Pass the summed values to the inputs
  document.getElementById('inputColD').value = sumColD;
  document.getElementById('inputColE').value = sumColEDate.toLocaleDateString(); // Format as short date
  document.getElementById('inputColH').value = sumColH;
}

// Function to handle the page load event and initialize the Excel connection
async function initializeExcelConnection(mainWindow) {

  workbook = new ExcelJS.Workbook();
  const filePath = 'D:\\xx\\db.xlsx'; // Path to your Excel file
  await workbook.xlsx.readFile(filePath);
  worksheet = workbook.getWorksheet('ProH'); // Default sheet (adjust for each page as needed)
  
  // Now loop through sheet rows from row 2 to the last row with data
  const matchedNodes = [];
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber >= 3) { // Start from row 3 (skipping the header)
      const col1Value = row.getCell(1).value;
      const col2Value = row.getCell(2).value;
 
      // Combine the values into a string "col1 - col2"
      matchedNodes.push(`${col1Value} - ${col2Value}`);
    }
  });

  // Update the treeview with the matched nodes
  mainWindow.webContents.send('update-treeview', matchedNodes);
}

module.exports = { initializeExcelConnection };

// Function to update the treeview with matched nodes
function updateTreeView(nodes) {
  const treeview = document.getElementById('treeview');
  const headNode = document.createElement('li');
  headNode.innerText = 'Projects';
  treeview.appendChild(headNode);

  nodes.forEach(node => {
    const subNode = document.createElement('li');
    subNode.innerText = node;
    subNode.onclick = nodeClicked; // Assign the node click handler
    headNode.appendChild(subNode);
  });
}



