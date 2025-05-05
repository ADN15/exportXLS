// Async function to load scripts
var getScriptPromisify = (src) => {
    return new Promise(resolve => {
        fetch(src)
            .then(response => response.text())
            .then(scriptText => {
                const script = document.createElement('script');
                script.textContent = scriptText;
                document.head.appendChild(script);
                resolve();
            })
            .catch(error => {
                console.error(`Failed to load script: ${src}`, error);
                resolve();
            });
    });
};

(function () {
    const template = document.createElement('template');
    template.innerHTML = `
    <style>
    button {
        padding: 5px 10px;
        background-color: #007bff;
        color: #fff;
        border: none;
        cursor: pointer;
    }
    </style>
    <section>
        <button id="exportButton">Export Excel</button>
    </section>
    `;

    class newExportXLS extends HTMLElement {
        constructor() {
            super();

            // HTML objects
            this.attachShadow({ mode: 'open' });
            this.shadowRoot.appendChild(template.content.cloneNode(true));
            this._exportButton = this.shadowRoot.querySelector('#exportButton');

            // Load SheetJS dynamically before binding the click event
            getScriptPromisify("https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js").then(() => {
                this._exportButton.addEventListener('click', () => this.exportData());
            });
        }

        // Method to export Excel data
        exportData(resultSet = []) {
            if (typeof XLSX === "undefined") {
                console.error("XLSX library not loaded.");
                return;
            }

            // Create a new workbook
            const wb = XLSX.utils.book_new();

            // Define the "Drawdown_Table" data
            //const ws_data = [];
            //ws_data.push(['', '', '', '', '', '', '', 'Measures', 'Revised-CFY', 'Estimated-NFY']);
            //ws_data.push(["Cost Centre", "Cost Centre Description", "Funding Pot", "Starting FY", "Closing FY", "Funding Pot Description", "Accounts", "Account Description", "", ""]);

            //resultSet.forEach(item => {
            //    const values = item.split(';');
            //    const rowData = [
            //        values[0], // Cost Centre
            //        values[1], // Cost Centre Description
            //        values[2], // Funding Pot
            //        values[3], // Starting FY
            //        values[4], // Closing FY
            //        values[5], // Funding Pot Description
            //        values[6], // Accounts
            //        values[7], // Account Description
            //        '',        // Revised-CFY (empty)
            //        ''         // Estimated-NFY (empty)
            //    ];
            //    ws_data.push(rowData);
            //});

            // Create the "Drawdown_Table" worksheet
            //const wsDrawdown = XLSX.utils.aoa_to_sheet(ws_data);

            // Protect the worksheet (making it read-only)
            //wsDrawdown["!protect"] = {
            //    password: "",  // Optional password (empty means no password)
            //    sheet: true,   // Lock the sheet
            //    formatCells: false,
            //    formatColumns: false,
            //    formatRows: false,
            //    insertColumns: false,
            //    insertRows: false,
            //    deleteColumns: false,
            //    deleteRows: false
            //};

            // Add the "Drawdown_Table" worksheet to the workbook
            //XLSX.utils.book_append_sheet(wb, wsDrawdown, "Drawdown_Table");


            // Create the "Date" sheet data
            const wsDateData = [];
            wsDateData.push(['Measures', 'Budget Allocation']);
            wsDateData.push(['Date', '']);
            resultSet.forEach(item => {
                const values = item.split(';');
                const rowData = [
                    '', // Year
                    ''  // Value
                ];
                wsDateData.push(rowData);
            });

            // Create the "Date" worksheet
            const wsDate = XLSX.utils.aoa_to_sheet(wsDateData);

            wsDate["!protect"] = {
                password: "",  // Optional password (empty means no password)
                sheet: true,   // Lock the sheet
                formatCells: false,
                formatColumns: false,
                formatRows: false,
                insertColumns: false,
                insertRows: false,
                deleteColumns: false,
                deleteRows: false
            };

            // Add the "Date" worksheet to the workbook
            XLSX.utils.book_append_sheet(wb, wsDate, "Date");

            // Create the "CostCenter" sheet data
            const wsCostCenterData = [];
            wsCostCenterData.push(['', '', 'Measures','Count']);
            wsCostCenterData.push(['Ministry View', 'Programme', 'Description','']);
            resultSet.forEach(item => {
                const values = item.split(';');
                const rowData = [
                    values[1],  // Cost Centre Description
                    values[0], // Cost Centre ID
                    '',
                    ''
                ];
                wsCostCenterData.push(rowData);
            });

            // Create the "CostCenter" worksheet
            const wsCostCenter = XLSX.utils.aoa_to_sheet(wsCostCenterData);

            wsCostCenter["!protect"] = {
                password: "",  // Optional password (empty means no password)
                sheet: true,   // Lock the sheet
                formatCells: false,
                formatColumns: false,
                formatRows: false,
                insertColumns: false,
                insertRows: false,
                deleteColumns: false,
                deleteRows: false
            };

            // Add the "CostCenter" worksheet to the workbook
            XLSX.utils.book_append_sheet(wb, wsCostCenter, "CostCenter");

            // Create the "FundingPot" sheet data
            const wsFundingPotData = [];
            wsFundingPotData.push(['', '', '', '', '','Measures','Total Supported']);
            wsFundingPotData.push(['', '', '', '', '','Version','Estimated']);
            wsFundingPotData.push(['Funding Pot', 'Allocated Access', 'Startin FY', 'Closeing FY', 'Description','Accounts','']);
            resultSet.forEach(item => {
                const values = item.split(';');
                const rowData = [
                    values[2], // Funding Pot ID
                    '',        // Allocated Access
                    values[3], // Starting FY
                    values[4], // Closing FY
                    values[5], // Funding Pot Description
                    '',        // Account
                    ''         // Total Supported Extimated (empty for now)
                ];
                wsFundingPotData.push(rowData);
            });

            // Create the "FundingPot" worksheet
            const wsFundingPot = XLSX.utils.aoa_to_sheet(wsFundingPotData);

            wsFundingPot["!protect"] = {
                password: "",  // Optional password (empty means no password)
                sheet: true,   // Lock the sheet
                formatCells: false,
                formatColumns: false,
                formatRows: false,
                insertColumns: false,
                insertRows: false,
                deleteColumns: false,
                deleteRows: false
            };

            // Add the "FundingPot" worksheet to the workbook
            XLSX.utils.book_append_sheet(wb, wsFundingPot, "FundingPot");

            // Create the "Account" sheet data
            const wsAccountData = [];
            wsAccountData.push(['', '', '','Measures','Count']);
            wsAccountData.push(['Accounts', 'Old or New Account', 'Old Long Decsription','Description','']);
            resultSet.forEach(item => {
                const values = item.split(';');
                const rowData = [
                    values[6], // Account
                    '',        // Old or New Account
                    '',        // Old Long Descriotipn
                    values[7], // Account Description
                    ''         // Old_New (empty for now)
                ];
                wsAccountData.push(rowData);
            });

            // Create the "Account" worksheet
            const wsAccount = XLSX.utils.aoa_to_sheet(wsAccountData);

            wsAccount["!protect"] = {
                password: "",  // Optional password (empty means no password)
                sheet: true,   // Lock the sheet
                formatCells: false,
                formatColumns: false,
                formatRows: false,
                insertColumns: false,
                insertRows: false,
                deleteColumns: false,
                deleteRows: false
            };

            // Add the "Account" worksheet to the workbook
            XLSX.utils.book_append_sheet(wb, wsAccount, "Account");

            // Generate Excel file and trigger download
            const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
            const blob = new Blob([wbout], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
            const filename = "Budget Drawdown Data.xlsx";

            const link = document.createElement("a");
            link.href = URL.createObjectURL(blob);
            link.download = filename;
            document.body.appendChild(link);
            link.click();

            // Cleanup
            setTimeout(() => {
                URL.revokeObjectURL(link.href);
                link.remove();
            }, 100);

            // Dispatch a custom event indicating successful export
            this.dispatchEvent(new CustomEvent('onFileExport', { detail: { filename } }));
        }


        // Standard Web Component function used to add event listeners
        connectedCallback() {
            // Additional setup if needed
        }
    }

    window.customElements.define('com-sap-new-version-export-xls', newExportXLS);
})();
