import * as XLSX from "xlsx";
import { saveAs } from 'file-saver';

import './App.css';

function App() {

  const dataArray = [
    "924027,Test 001,A_240006,FY2023,FY2025,Test 001,T5111,Manpower,1000,6000",
    "924027,Test 001,A_240006,FY2023,FY2025,Test 001,T5111,Manpower,1000,6000"
  ];

  // Function to export dataArray data to XLS format
  const exportExcel = (dataArray) => {
    // Convert array data into array of arrays
    const data = dataArray.map(item => {
      const [costCenter, costCenterDesc, fundingPot, startingFY, closingFY, fundingPorDesc, account, accountDesc, revisedCFY, estimatedNFY] = item.split(',');
      return [costCenter, costCenterDesc, fundingPot, startingFY, closingFY, fundingPorDesc, account, accountDesc, revisedCFY, estimatedNFY];
    });

    // Headers
    const headers1 = ['', '', '', '', '', '', '', 'Measures', 'Revised-CFY' ,'Estimated-NFY'];
    const headers2 = ["Cost Center","Cost Center Description","Funding Pot", "Starting FY", "Closing FY","Funding Pot Description","Accounts","Account Description","",""];
    data.unshift(headers1, headers2);

    // Convert the array data to a worksheet
    const ws = XLSX.utils.aoa_to_sheet(data);

    // Apply cell protection to rows 1 and 2
    const lockedRows = [1, 2];
    const headers = data[0];
    for (let r = 0; r < data.length; r++) {
    for (let c = 0; c < headers.length; c++) {
        const cell = XLSX.utils.encode_cell({ r: r, c: c });
        if (lockedRows.includes(r)) {
        // Apply protection for rows 1 and 2
        if (!ws[cell]) ws[cell] = {};
        ws[cell].s = { protection: { locked: true } };
        } else {
        // Set protection to unlocked for other rows
        if (!ws[cell]) ws[cell] = {};
        ws[cell].s = { protection: { locked: false } };
        }
    }
    }

    // Apply cell styling
    const headerCellStyle = {
      font: { bold: true },
      alignment: { horizontal: 'center' },
      fill: { fgColor: { rgb: "FFCC00" } } // Light orange color code
    };
    const cellStyle = {
      alignment: { horizontal: 'center' }
    };

    // Apply styling to headers
    headers.forEach((header, index) => {
      const cell = XLSX.utils.encode_cell({ r: 0, c: index });
      ws[cell] = { ...ws[cell], ...headerCellStyle };
    });

    // Apply styling to cells
    for (let r = 1; r <= data.length; r++) {
      for (let c = 0; c < headers.length; c++) {
        const cell = XLSX.utils.encode_cell({ r: r, c: c });
        ws[cell] = { ...ws[cell], ...cellStyle };
      }
    }

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
