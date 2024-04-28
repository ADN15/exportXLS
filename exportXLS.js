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

class ExportXLS extends HTMLElement {
    constructor() {
        super();

        // HTML objects
        this.attachShadow({ mode: 'open' });
        this.shadowRoot.appendChild(template.content.cloneNode(true));
        this._exportButton = this.shadowRoot.querySelector('#exportButton');

        // Binding the click event for export
        this._exportButton.addEventListener('click', () => this.exportData());
    }
    // Method to export Excel data
    exportData(resultSet) {
        // Prepare the Excel content
        let excelContent = '';

        // Add headers
        const headers1 = ['', '', '', '', '', '', '', 'Measures', 'Revised-CFY', 'Estimated-NFY'];
        const headers2 = ["Cost Centre", "Cost Centre Description", "Funding Pot", "Starting FY", "Closing FY", "Funding Pot Description", "Accounts", "Account Description", "", ""];
        excelContent += headers1.join('\t') + '\n';
        excelContent += headers2.join('\t') + '\n';

        // Add data
        resultSet.forEach(item => {
            const values = item.split(';');
            const [costCenter, costCenterDesc, fundingPot, startingFY, closingFY, fundingPotDesc, account, accountDesc] = values;
            // Extract specific values from the string and map them to the corresponding headers
            const rowData = [
                costCenter,                 // Cost Center
                costCenterDesc,             // Cost Center Description
                fundingPot,                 // Funding Pot
                startingFY,                 // Starting FY
                closingFY,                  // Closing FY
                fundingPotDesc,            // Funding Pot Description
                account,                    // Accounts
                accountDesc,                // Account Description
                '',                         // Revised-CFY (empty for now)
                ''                          // Estimated-NFY (empty for now)
            ];
            excelContent += rowData.map(cell => '"' + cell.replace(/"/g, '""') + '"').join('\t') + '\n';
        });

        // Modify the default sheet name while generating content
        console.log("excel content");
        console.log(excelContent);

        // Create a proper byte array with UTF-16LE encoding
        const utf16Array = new Uint16Array(excelContent.length);
        for (let i = 0; i < excelContent.length; i++) {
            utf16Array[i] = excelContent.charCodeAt(i);
        }

        // Convert the byte array to Blob
        const blob = new Blob([new Uint8Array([0xFF, 0xFE]), utf16Array], { type: 'application/vnd.ms-excel' });

        // Generate a filename with timestamp
        const filename = `Budget Drawdown.xls`;
        // Set the sheet name to a specific value
        const sheetname = 'Drawdown_Table';

        // Create a temporary anchor element to trigger the download
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = filename;
        link.setAttribute('sheetname', sheetname);

        // Dispatch a custom event indicating successful export
        //this.dispatchEvent(new CustomEvent('onFileExport', { detail: filename }));
        this.dispatchEvent(new CustomEvent('onFileExport', { detail: { filename, sheetname } }));


        // Trigger the download
        link.click();

        // Cleanup
        setTimeout(() => {
            URL.revokeObjectURL(link.href);
            link.remove();
        }, 100);
    }

    // Standard Web Component function used to add event listeners
    connectedCallback() {
        // Add any additional setup or event listeners here if needed
    }
}

window.customElements.define('com-sap-sample-export-xls', ExportXLS);
})();
