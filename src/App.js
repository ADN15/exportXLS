import * as XLSX from "xlsx";
import { saveAs } from 'file-saver';
import './App.css';

function App() {
  const dataArray = [
    "924150;Test 006;W_250002;FY2025;FY2028;B-A-01 Test 002;5113;Purchase of Tangible Assets - Plants and Machinery;;1000",
    "924150;Test 006;W_250002;FY2025;FY2028;B-A-01 Test 002;5113;Purchase of Tangible Assets - Plants and Machinery;;1000"
  ];

  // Function to export dataArray data to XLS format
  const exportExcel = async (dataArray) => {
    // Convert array data into array of arrays
    const data = dataArray.map(item => item.split(';'));

    // Headers
    const headers1 = ['', '', '', '', '', '', '', 'Measures', 'Revised-CFY' ,'Estimated-NFY'];
    const headers2 = ["Cost Center","Cost Center Description","Funding Pot", "Starting FY", "Closing FY","Funding Pot Description","Accounts","Account Description","",""];
    data.unshift(headers1, headers2);

    // Convert the array data to a worksheet
    const ws = XLSX.utils.aoa_to_sheet(data);

    // Apply cell protection to specific columns
    const lockedColumns = [1, 2]; // Indexes of columns to lock
    const headers = data[0];
    const secondRow = data[1];
    for (let c = 0; c < headers.length; c++) {
      if (lockedColumns.includes(c)) {
        // Apply protection for specified columns based on values in the first and second rows
        const shouldLock = secondRow[c] !== '' && headers[c] !== ''; // Lock if both rows have values
        for (let r = 0; r < data.length; r++) {
          const cell = XLSX.utils.encode_cell({ r: r, c: c });
          if (!ws[cell]) ws[cell] = {};
          ws[cell].s = { protection: { locked: shouldLock, lockText: true } };
        }
      } else {
        // If the column index is not in lockedColumns, unlock the column
        for (let r = 0; r < data.length; r++) {
          const cell = XLSX.utils.encode_cell({ r: r, c: c });
          if (!ws[cell]) ws[cell] = {};
          ws[cell].s = { protection: { locked: false, lockText: true } };
        }
      }
    }

    // Format the cells as numbers and add a separator for the last two columns
    const lastColumnIndex = headers.length - 1;
    const secondLastColumnIndex = lastColumnIndex - 1;
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let r = range.s.r; r <= range.e.r; r++) {
      for (let c = secondLastColumnIndex; c <= lastColumnIndex; c++) {
        const cell = XLSX.utils.encode_cell({ r: r, c: c });
        const cellValue = ws[cell] ? ws[cell].v : '';
        if (!isNaN(cellValue)) {
          ws[cell].t = 'n'; // Set cell type to number
          ws[cell].z = '#,##0'; // Set number format with comma separator
        }
      }
    }

    // Set sheet protection
    ws['!protect'] = {
        selectLockedCells: true,
        selectUnlockedCells: true,
        formatCells: false,
        formatColumns: false,
        formatRows: false,
        insertRows: false,
        insertColumns: false,
        insertHyperlinks: false,
        deleteRows: false,
        deleteColumns: false,
        sort: false,
        autoFilter: false,
        pivotTables: false,
        password: 'password'
    };

    // Create a new workbook and add the worksheet
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

    // Generate XLSX file buffer
    const buffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });

    // Create Blob from buffer
    const blob = new Blob([buffer], { type: 'application/octet-stream' });

    // Trigger file download
    const fileName = "custom_data.xlsx";
    saveAs(blob, fileName);
  }

  return (
    <div className="App">
      <input 
        type="button"
        value="Export to Excel"
        onClick={() => exportExcel(dataArray)} 
      />
    </div>
  );
}

export default App;
