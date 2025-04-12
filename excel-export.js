async function exportToExcel() {
    try {
      console.log("Starting export...");
  
      const dashboard = tableau.extensions.dashboardContent.dashboard;
      const worksheet = dashboard.worksheets[0];
      const summaryData = await worksheet.getSummaryDataAsync();
  
      console.log("Summary data retrieved:", summaryData);
  
      const data = summaryData.data;
      const columns = summaryData.columns;
  
      const rowFields = getSelectedValues("rowFields");
      const colFields = getSelectedValues("columnFields");
      const valueFields = getSelectedValues("valueFields");
  
      console.log("Row fields:", rowFields);
      console.log("Column fields:", colFields);
      console.log("Value fields:", valueFields);
  
      if (valueFields.length === 0) {
        alert("Please select at least one value field.");
        return;
      }
  
      const title = document.getElementById("title").value;
      const fontSize = parseInt(document.getElementById("headerFontSize").value);
      const isBold = document.getElementById("headerBold").checked;
  
      const conditionType = document.getElementById("conditionType").value;
      const conditionValue = parseFloat(document.getElementById("conditionValue").value);
      const highlightColor = document.getElementById("highlightColor").value;
  
      console.log("Export settings:", { title, fontSize, isBold, conditionType, conditionValue, highlightColor });
  
      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet("Export");
  
      // Add title row
      if (title) {
        console.log("Adding title to worksheet:", title);
        sheet.addRow([title]);
        sheet.mergeCells(`A1:${String.fromCharCode(65 + columns.length - 1)}1`);
        const titleRow = sheet.getRow(1);
        titleRow.font = { size: fontSize + 2, bold: true };
        titleRow.alignment = { horizontal: "center" };
      }
  
      // Add the headers row (raw data columns)
      const headerRow = sheet.addRow(["", ...columns.map(col => col.fieldName)]);
      headerRow.font = { size: fontSize, bold: isBold };
      console.log("Added header row:", headerRow.values);
  
      // Add the raw data
      data.forEach((row, rowIndex) => {
        const excelRow = sheet.addRow(row.map(cell => cell.formattedValue || ""));
        console.log("Added row to Excel:", excelRow.values);
  
        // Apply conditional formatting if needed
        if (conditionType && conditionValue) {
          excelRow.eachCell((cell, colNumber) => {
            const value = cell.value;
            if (conditionType === "greaterThan" && value > conditionValue) {
              cell.style = { fill: { type: "pattern", pattern: "solid", fgColor: { argb: highlightColor.replace('#', '') } } };
            } else if (conditionType === "lessThan" && value < conditionValue) {
              cell.style = { fill: { type: "pattern", pattern: "solid", fgColor: { argb: highlightColor.replace('#', '') } } };
            }
          });
        }
      });
  
      // Finalize the file
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const link = document.createElement("a");
      link.href = URL.createObjectURL(blob);
      link.download = "tableau_export.xlsx";
      link.click();
  
      console.log("Export successful!");
    } catch (error) {
      console.error("Error exporting to Excel: ", error);
      alert("An error occurred while exporting to Excel.");
    }
  }
  