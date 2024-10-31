let excelData = []; // Placeholder for Excel data
let filteredData = []; // Placeholder for filtered data

// Load the Google Sheets file when the page loads
document.addEventListener('DOMContentLoaded', async () => {
    const urlParams = new URLSearchParams(window.location.search);
    const fileUrl = urlParams.get('fileUrl');

    if (fileUrl) {
        await loadExcelData(fileUrl);
    }
});

// Function to load Excel data
async function loadExcelData(url) {
    const response = await fetch(url);
    const data = await response.arrayBuffer();
    const workbook = XLSX.read(data);
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    excelData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
    displayData(excelData);
}

// Function to display data in the table
function displayData(data) {
    const sheetContent = document.getElementById('sheet-content');
    sheetContent.innerHTML = '';

    const table = document.createElement('table');
    data.forEach((row, rowIndex) => {
        const tr = document.createElement('tr');
        row.forEach((cell, cellIndex) => {
            const td = document.createElement('td');
            td.textContent = cell;
            td.addEventListener('click', () => toggleHighlight(rowIndex, cellIndex));
            tr.appendChild(td);
        });
        table.appendChild(tr);
    });
    sheetContent.appendChild(table);
}

// Array to track highlighted cells
let highlightedCells = [];

// Toggle highlight on cell click
function toggleHighlight(row, col) {
    const cellIndex = `${row}-${col}`;
    if (highlightedCells.includes(cellIndex)) {
        highlightedCells = highlightedCells.filter(cell => cell !== cellIndex);
    } else {
        highlightedCells.push(cellIndex);
    }
    highlightData();
}

// Highlight selected cells
function highlightData() {
    const rows = document.querySelectorAll('#sheet-content tr');
    rows.forEach((tr, rowIndex) => {
        const cells = tr.querySelectorAll('td');
        cells.forEach((td, colIndex) => {
            const cellIndex = `${rowIndex}-${colIndex}`;
            if (highlightedCells.includes(cellIndex)) {
                td.style.backgroundColor = '#ffeb3b'; // Highlight color
            } else {
                td.style.backgroundColor = ''; // Remove highlight
            }
        });
    });
}

// Apply operation on selected data
document.getElementById('apply-operation').addEventListener('click', () => {
    const rowFrom = parseInt(document.getElementById('row-range-from').value) - 1;
    const rowTo = parseInt(document.getElementById('row-range-to').value) - 1;
    const operationType = document.getElementById('operation-type').value;
    const operation = document.getElementById('operation').value;

    // Filter data based on the selected rows and operation
    filteredData = excelData.filter((_, rowIndex) => rowIndex >= rowFrom && rowIndex <= rowTo);
    displayData(filteredData); // Show filtered data
});

// Download functionality
document.getElementById('download-button').addEventListener('click', () => {
    const modal = document.getElementById('download-modal');
    modal.style.display = 'flex';
});

// Close modal
document.getElementById('close-modal').addEventListener('click', () => {
    document.getElementById('download-modal').style.display = 'none';
});

// Confirm download
document.getElementById('confirm-download').addEventListener('click', () => {
    const filename = document.getElementById('filename').value;
    const format = document.getElementById('file-format').value;
    if (format === 'xlsx') {
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.json_to_sheet(filteredData);
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
        XLSX.writeFile(wb, `${filename}.xlsx`);
    } else {
        const csvData = filteredData.map(row => row.join(',')).join('\n');
        const blob = new Blob([csvData], { type: 'text/csv' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = `${filename}.csv`;
        link.click();
    }
});
