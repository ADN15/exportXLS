import { useState } from "react";
import * as XLSX from "xlsx";
import { saveAs } from 'file-saver';

import './App.css';

function App() {
  const [parsedDataString, setParsedDataString] = useState("");

  const resultSet = [
    {
        "Date": {
            "id": "[Date].[YM].[Date.YEAR].[2024]",
            "description": "2024",
            "parentId": "[Date].[YM].[All].[(all)]",
            "properties": {}
        },
        "Version": {
            "id": "public.Actual",
            "description": "Actual",
            "properties": {}
        },
        "@MeasureDimension": {
            "id": "Revenue",
            "description": "Revenue",
            "rawValue": "1000",
            "formattedValue": "1,000.00"
        },
        "Company_Code": {
            "id": "SG01",
            "description": "Singapore 1",
            "properties": {}
        },
        "Rule": {
            "id": "N",
            "description": "No",
            "properties": {}
        }
    },
    {
        "Date": {
            "id": "[Date].[YM].&[202401]",
            "description": "Jan (2024)",
            "parentId": "[Date].[YM].[Date.YEAR].[2024]",
            "properties": {}
        },
        "Version": {
            "id": "public.Actual",
            "description": "Actual",
            "properties": {}
        },
        "@MeasureDimension": {
            "id": "Revenue",
            "description": "Revenue",
            "rawValue": "1000",
            "formattedValue": "1,000.00"
        },
        "Company_Code": {
            "id": "SG01",
            "description": "Singapore 1",
            "properties": {}
        },
        "Rule": {
            "id": "N",
            "description": "No",
            "properties": {}
        }
    },
    {
        "Date": {
            "id": "[Date].[YM].[Date.YEAR].[2024]",
            "description": "2024",
            "parentId": "[Date].[YM].[All].[(all)]",
            "properties": {}
        },
        "Version": {
            "id": "public.Actual",
            "description": "Actual",
            "properties": {}
        },
        "@MeasureDimension": {
            "id": "Revenue",
            "description": "Revenue",
            "rawValue": "4000",
            "formattedValue": "4,000.00"
        },
        "Company_Code": {
            "id": "SG01",
            "description": "Singapore 1",
            "properties": {}
        },
        "Rule": {
            "id": "Y",
            "description": "Yes",
            "properties": {}
        }
    },
    {
        "Date": {
            "id": "[Date].[YM].&[202401]",
            "description": "Jan (2024)",
            "parentId": "[Date].[YM].[Date.YEAR].[2024]",
            "properties": {}
        },
        "Version": {
            "id": "public.Actual",
            "description": "Actual",
            "properties": {}
        },
        "@MeasureDimension": {
            "id": "Revenue",
            "description": "Revenue",
            "rawValue": "4000",
            "formattedValue": "4,000.00"
        },
        "Company_Code": {
            "id": "SG01",
            "description": "Singapore 1",
            "properties": {}
        },
        "Rule": {
            "id": "Y",
            "description": "Yes",
            "properties": {}
        }
    }
];

  // Function to export resultSet data to XLS format
  const exportExcel = (resultSet) => {
    // Transform resultSet data into pivot-style format
    const pivotData = {};
  
    resultSet.forEach(item => {
        const { Date, Version, "@MeasureDimension": measureDimension, Company_Code, Rule } = item;
        const key = `${Company_Code.description}_${Rule.description}`;
  
        if (!pivotData[key]) {
            pivotData[key] = {};
        }
  
        const dateKey = `${Date.description}_${Version.description}`;
        pivotData[key][dateKey] = parseFloat(measureDimension.rawValue); // Use parseFloat to convert string to number
    });
  
    // Convert pivot data into array of arrays
    const data = [];
  
    // Get unique dates and versions
    const dates = new Set();
    const versions = new Set();
  
    for (const key in pivotData) {
        for (const dateKey in pivotData[key]) {
            const [date, version] = dateKey.split('_');
            dates.add(date);
            versions.add(version);
        }
    }
  
    // Headers
    const headers = ['', 'Date', 'Version', ...Array.from(versions)];
  
    data.push(headers);
  
    // Rows
    for (const key in pivotData) {
        const [companyCode, rule] = key.split('_');
        const row = [companyCode, rule];
  
        dates.forEach(date => {
            versions.forEach(version => {
                const value = pivotData[key][`${date}_${version}`] || '';
                row.push(value);
            });
        });
  
        data.push(row);
    }
  
    // Convert the array data to a worksheet
    const ws = XLSX.utils.aoa_to_sheet(data);
  
    // Apply cell styling
    const yellowCellStyle = {
      fill: {
        fgColor: { rgb: "FFFF00" } // Yellow color code
      }
    };

    // Determine column index for Revenue
    const revenueColumnIndex = headers.findIndex(header => header === '@MeasureDimension');
    console.log("Revenue Column Index:", revenueColumnIndex);

    // Loop through each row to check Revenue value and apply styling
    for (let r = 2; r <= data.length; r++) {
      const cell = XLSX.utils.encode_cell({ r: r, c: revenueColumnIndex });
      if (ws[cell] && parseFloat(ws[cell].v) >= 1000) { // Check if Revenue is more or equal to 1000
        console.log(parseFloat(ws[cell].v));
        ws[cell].s = yellowCellStyle;
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
        value="Test"
        onClick={() => exportExcel(resultSet)} 
      />
    </div>
  );
}

export default App;
