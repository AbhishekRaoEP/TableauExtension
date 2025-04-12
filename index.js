document.addEventListener("DOMContentLoaded", () => {
    tableau.extensions.initializeAsync().then(() => {
      populateFieldSelectors();
      // Ensure we're using the exportToExcel from excel-export.js
      document.getElementById("exportBtn").addEventListener("click", exportToExcel);
    });
  });
  
  async function populateFieldSelectors() {
    const worksheet = tableau.extensions.dashboardContent.dashboard.worksheets[0];
    const summaryData = await worksheet.getSummaryDataAsync();
    const fields = summaryData.columns.map(col => col.fieldName);
  
    const rowSelect = document.getElementById("rowFields");
    const colSelect = document.getElementById("columnFields");
    const valSelect = document.getElementById("valueFields");
  
    [rowSelect, colSelect, valSelect].forEach(select => {
      select.innerHTML = "";
      fields.forEach(field => {
        const option = document.createElement("option");
        option.value = field;
        option.textContent = field;
        select.appendChild(option.cloneNode(true));
      });
    });
  }
  