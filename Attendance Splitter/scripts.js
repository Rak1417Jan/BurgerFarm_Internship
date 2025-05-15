// document.addEventListener("DOMContentLoaded", function () {
//   const fileInput = document.getElementById("excel-file");
//   const fileName = document.getElementById("file-name");
//   const progressContainer = document.getElementById("progress-container");
//   const progressBar = document.getElementById("progress-bar");
//   const alertContainer = document.getElementById("alert-container");
//   const loader = document.getElementById("loader");
//   const resultsContainer = document.getElementById("results-container");
//   const branchesList = document.getElementById("branches-list");
//   const downloadAllBtn = document.getElementById("download-all-btn");

//   let branchData = {};
//   let originalHeaders = [];
//   let workbook = null;

//   // Handle file selection
//   fileInput.addEventListener("change", function (e) {
//     const file = e.target.files[0];

//     if (!file) return;

//     // Display file name
//     fileName.textContent = file.name;
//     fileName.style.display = "block";

//     // Clear previous results
//     branchData = {};
//     branchesList.innerHTML = "";
//     resultsContainer.style.display = "none";
//     alertContainer.innerHTML = "";

//     // Show progress
//     progressContainer.style.display = "block";
//     progressBar.style.width = "0%";

//     // Animate progress bar
//     let progress = 0;
//     const progressInterval = setInterval(() => {
//       progress += 5;
//       progressBar.style.width = `${Math.min(progress, 90)}%`;
//       if (progress >= 90) clearInterval(progressInterval);
//     }, 100);

//     // Process the file
//     processExcelFile(file)
//       .then(() => {
//         // Complete progress bar
//         clearInterval(progressInterval);
//         progressBar.style.width = "100%";

//         // Hide loader and show results after a short delay
//         setTimeout(() => {
//           loader.style.display = "none";
//           resultsContainer.style.display = "block";
//           showAlert(
//             "File processed successfully! You can now download files by branch.",
//             "success"
//           );
//         }, 500);
//       })
//       .catch((error) => {
//         clearInterval(progressInterval);
//         progressBar.style.width = "0%";
//         progressContainer.style.display = "none";
//         loader.style.display = "none";
//         showAlert(error.message, "danger");
//       });
//   });

// async function processExcelFile(file) {
//     return new Promise((resolve, reject) => {
//         const reader = new FileReader();

//         reader.onload = function (e) {
//             try {
//                 loader.style.display = "block";

//                 // Parse workbook
//                 const data = new Uint8Array(e.target.result);
//                 workbook = XLSX.read(data, { type: "array" });

//                 // Get first sheet
//                 const firstSheetName = workbook.SheetNames[0];
//                 const worksheet = workbook.Sheets[firstSheetName];

//                 // Convert to JSON
//                 const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

//                 // Check if file has data
//                 if (jsonData.length < 2) {
//                     throw new Error(
//                         "The Excel file seems to be empty or has insufficient data."
//                     );
//                 }

//                 // Get headers
//                 originalHeaders = jsonData[0];

//                 // Find the critical column indices
//                 const branchIndex = originalHeaders.findIndex(
//                     (header) =>
//                         header && header.toString().toLowerCase().trim() === "branch"
//                 );
//                 const nameIndex = originalHeaders.findIndex(
//                     (header) =>
//                         header && header.toString().toLowerCase().trim() === "name"
//                 );
//                 const idIndex = originalHeaders.findIndex(
//                     (header) =>
//                         header && header.toString().toLowerCase().trim() === "id"
//                 );

//                 if (branchIndex === -1) {
//                     throw new Error(
//                         'Could not find a "Branch" column in the Excel file. Please ensure the file contains this column.'
//                     );
//                 }

//                 // Group data by branch
//                 for (let i = 1; i < jsonData.length; i++) {
//                     const currentRow = jsonData[i];
//                     if (!currentRow || currentRow.length <= branchIndex) continue;

//                     // Check if this is an employee header row (contains actual branch info)
//                     const isHeaderRow = currentRow.some((cell, index) => {
//                         // Skip branch, name, id columns from this check
//                         if (index === branchIndex || index === nameIndex || index === idIndex) {
//                             return false;
//                         }
//                         // If any cell in this row has non-time data (like "OPERATION Shift")
//                         return cell && !isTimeValue(cell.toString().trim());
//                     });

//                     if (isHeaderRow) {
//                         // This is an employee header row
//                         const branch = currentRow[branchIndex]
//                             ? currentRow[branchIndex].toString().trim()
//                             : "Undefined";

//                         const employeeName = nameIndex !== -1 && currentRow[nameIndex]
//                             ? currentRow[nameIndex].toString().trim()
//                             : "Unknown";

//                         const employeeId = idIndex !== -1 && currentRow[idIndex]
//                             ? currentRow[idIndex].toString().trim()
//                             : "";

//                         // Create branch entry if it doesn't exist
//                         if (!branchData[branch]) {
//                             branchData[branch] = [originalHeaders];
//                         }

//                         // Add the header row
//                         branchData[branch].push(currentRow);

//                         // Add the next 6 rows (time entries) for this employee
//                         for (let j = 1; j <= 6 && (i + j) < jsonData.length; j++) {
//                             const timeRow = jsonData[i + j];
//                             if (timeRow && timeRow.length > 0) {
//                                 // Copy the identifying information from the header row
//                                 if (idIndex !== -1) timeRow[idIndex] = employeeId;
//                                 if (nameIndex !== -1) timeRow[nameIndex] = employeeName;
//                                 if (branchIndex !== -1) timeRow[branchIndex] = branch;
//                                 branchData[branch].push(timeRow);
//                             }
//                         }

//                         // Skip the next 6 rows as we've already processed them
//                         i += 6;
//                     }
//                 }

//                 // Display branch information
//                 displayBranchResults();

//                 resolve();
//             } catch (error) {
//                 reject(error);
//             }
//         };

//         reader.onerror = function () {
//             reject(new Error("Failed to read the file. Please try again."));
//         };

//         reader.readAsArrayBuffer(file);
//     });
// }

// // Helper function to check if a value represents time data
// function isTimeValue(value) {
//     if (!value) return true; // Empty cells are allowed in time rows
//     if (value.toLowerCase() === "in" || value.toLowerCase() === "out" || 
//         value.toLowerCase() === "hours" || value.toLowerCase() === "ot") {
//         return true;
//     }
//     // Check for decimal numbers (Excel time format)
//     if (!isNaN(value) && value.toString().includes('.')) {
//         return true;
//     }
//     return false;
// }

//   // Display branch results
//   function displayBranchResults() {
//     branchesList.innerHTML = "";

//     // Sort branches alphabetically
//     const sortedBranches = Object.keys(branchData).sort();

//     sortedBranches.forEach((branch) => {
//       const rowCount = branchData[branch].length - 1; // Subtract header row
//       const employeeCount = Math.floor(rowCount / 7); // Each employee has 7 rows

//       const branchItem = document.createElement("div");
//       branchItem.className = "branch-item";

//       const branchName = document.createElement("div");
//       branchName.className = "branch-name";
//       branchName.textContent = branch;

//       const branchCount = document.createElement("div");
//       branchCount.className = "branch-count";
//       branchCount.textContent = `${employeeCount} employees (${rowCount} records)`;

//       const downloadBtn = document.createElement("a");
//       downloadBtn.className = "download-btn";
//       downloadBtn.href = "#";
//       downloadBtn.innerHTML = "<span>📥</span> Download";
//       downloadBtn.addEventListener("click", function (e) {
//         e.preventDefault();
//         downloadBranchExcel(branch);
//       });

//       branchItem.appendChild(branchName);
//       branchItem.appendChild(branchCount);
//       branchItem.appendChild(downloadBtn);

//       branchesList.appendChild(branchItem);
//     });

//     // Setup download all functionality
//     downloadAllBtn.addEventListener("click", function (e) {
//       e.preventDefault();
//       downloadAllBranches();
//     });
//   }

//   // Download Excel file for a specific branch
//   function downloadBranchExcel(branch) {
//     try {
//       // Create a new workbook
//       const newWorkbook = XLSX.utils.book_new();

//       // Convert branch data to worksheet
//       const worksheet = XLSX.utils.aoa_to_sheet(branchData[branch]);

//       // Add worksheet to workbook
//       XLSX.utils.book_append_sheet(newWorkbook, worksheet, "Branch Data");

//       // Generate Excel file
//       const excelBuffer = XLSX.write(newWorkbook, {
//         bookType: "xlsx",
//         type: "array",
//       });

//       // Save file
//       const blob = new Blob([excelBuffer], {
//         type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
//       });
//       saveAs(blob, `Branch_${branch}.xlsx`);

//       showAlert(
//         `Successfully downloaded data for branch: ${branch}`,
//         "success"
//       );
//     } catch (error) {
//       showAlert(`Failed to download: ${error.message}`, "danger");
//     }
//   }

//   // Download all branches as a zip file
//   async function downloadAllBranches() {
//     try {
//       const zip = new JSZip();

//       // Add each branch as a separate file
//       Object.keys(branchData).forEach((branch) => {
//         // Create a new workbook
//         const newWorkbook = XLSX.utils.book_new();

//         // Convert branch data to worksheet
//         const worksheet = XLSX.utils.aoa_to_sheet(branchData[branch]);

//         // Add worksheet to workbook
//         XLSX.utils.book_append_sheet(newWorkbook, worksheet, "Branch Data");

//         // Generate Excel file
//         const excelBuffer = XLSX.write(newWorkbook, {
//           bookType: "xlsx",
//           type: "array",
//         });

//         // Add to zip
//         zip.file(`Branch_${branch}.xlsx`, excelBuffer);
//       });

//       // Generate zip file
//       const content = await zip.generateAsync({ type: "blob" });

//       // Save zip file
//       saveAs(content, "All_Branches.zip");

//       showAlert(
//         "Successfully downloaded all branches as a zip file",
//         "success"
//       );
//     } catch (error) {
//       showAlert(`Failed to download all branches: ${error.message}`, "danger");
//     }
//   }

//   // Show alert message
//   function showAlert(message, type) {
//     const alert = document.createElement("div");
//     alert.className = `alert alert-${type}`;
//     alert.textContent = message;

//     alertContainer.innerHTML = "";
//     alertContainer.appendChild(alert);

//     // Auto hide after 5 seconds
//     setTimeout(() => {
//       alert.style.opacity = "0";
//       alert.style.transition = "opacity 0.5s";
//       setTimeout(() => {
//         if (alertContainer.contains(alert)) {
//           alertContainer.removeChild(alert);
//         }
//       }, 500);
//     }, 5000);
//   }
// });

document.addEventListener("DOMContentLoaded", function () {
  const fileInput = document.getElementById("excel-file");
  const fileName = document.getElementById("file-name");
  const progressContainer = document.getElementById("progress-container");
  const progressBar = document.getElementById("progress-bar");
  const alertContainer = document.getElementById("alert-container");
  const loader = document.getElementById("loader");
  const resultsContainer = document.getElementById("results-container");
  const branchesList = document.getElementById("branches-list");
  const downloadAllBtn = document.getElementById("download-all-btn");

  let branchData = {};
  let originalHeaders = [];
  let workbook = null;

  // Handle file selection
  fileInput.addEventListener("change", function (e) {
    const file = e.target.files[0];

    if (!file) return;

    // Display file name
    fileName.textContent = file.name;
    fileName.style.display = "block";

    // Clear previous results
    branchData = {};
    branchesList.innerHTML = "";
    resultsContainer.style.display = "none";
    alertContainer.innerHTML = "";

    // Show progress
    progressContainer.style.display = "block";
    progressBar.style.width = "0%";

    // Animate progress bar
    let progress = 0;
    const progressInterval = setInterval(() => {
      progress += 5;
      progressBar.style.width = `${Math.min(progress, 90)}%`;
      if (progress >= 90) clearInterval(progressInterval);
    }, 100);

    // Process the file
    processExcelFile(file)
      .then(() => {
        // Complete progress bar
        clearInterval(progressInterval);
        progressBar.style.width = "100%";

        // Hide loader and show results after a short delay
        setTimeout(() => {
          loader.style.display = "none";
          resultsContainer.style.display = "block";
          showAlert(
            "File processed successfully! You can now download files by branch.",
            "success"
          );
        }, 500);
      })
      .catch((error) => {
        clearInterval(progressInterval);
        progressBar.style.width = "0%";
        progressContainer.style.display = "none";
        loader.style.display = "none";
        showAlert(error.message, "danger");
      });
  });

  async function processExcelFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();

      reader.onload = function (e) {
        try {
          loader.style.display = "block";

          // Parse workbook
          const data = new Uint8Array(e.target.result);
          workbook = XLSX.read(data, { type: "array" });

          // Get first sheet
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];

          // Convert to JSON
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

          // Check if file has data
          if (jsonData.length < 2) {
            throw new Error(
              "The Excel file seems to be empty or has insufficient data."
            );
          }

          // Get headers
          originalHeaders = jsonData[0];

          // Find the critical column indices
          const branchIndex = originalHeaders.findIndex(
            (header) =>
              header && header.toString().toLowerCase().trim() === "branch"
          );
          const nameIndex = originalHeaders.findIndex(
            (header) =>
              header && header.toString().toLowerCase().trim() === "name"
          );
          const idIndex = originalHeaders.findIndex(
            (header) =>
              header && header.toString().toLowerCase().trim() === "id"
          );
          const departmentIndex = originalHeaders.findIndex(
            (header) =>
              header && header.toString().toLowerCase().trim() === "department"
          );

          if (branchIndex === -1) {
            throw new Error(
              'Could not find a "Branch" column in the Excel file. Please ensure the file contains this column.'
            );
          }

          // Group data by branch
          for (let i = 1; i < jsonData.length; i++) {
            const currentRow = jsonData[i];
            if (!currentRow || currentRow.length <= branchIndex) continue;

            // Check if this is an employee header row (contains actual branch info)
            const isHeaderRow = currentRow.some((cell, index) => {
              // Skip branch, name, id columns from this check
              if (index === branchIndex || index === nameIndex || index === idIndex || index === departmentIndex) {
                return false;
              }
              // If any cell in this row has non-time data (like "OPERATION Shift")
              return cell && !isTimeValue(cell.toString().trim());
            });

            if (isHeaderRow) {
              // This is an employee header row
              const branch = currentRow[branchIndex]
                ? currentRow[branchIndex].toString().trim()
                : "Undefined";

              const employeeName = nameIndex !== -1 && currentRow[nameIndex]
                ? currentRow[nameIndex].toString().trim()
                : "Unknown";

              const employeeId = idIndex !== -1 && currentRow[idIndex]
                ? currentRow[idIndex].toString().trim()
                : "";
                
              const department = departmentIndex !== -1 && currentRow[departmentIndex]
                ? currentRow[departmentIndex].toString().trim()
                : "";

              // Create branch entry if it doesn't exist
              if (!branchData[branch]) {
                branchData[branch] = [];
                // Add headers only once per branch
                branchData[branch].push(originalHeaders);
              }

              // Store employee data rows
              const employeeRows = [currentRow];

              // Add the next 6 rows (time entries) for this employee
              for (let j = 1; j <= 6 && (i + j) < jsonData.length; j++) {
                const timeRow = jsonData[i + j];
                if (timeRow && timeRow.length > 0) {
                  // Copy the identifying information from the header row
                  if (idIndex !== -1) timeRow[idIndex] = employeeId;
                  if (nameIndex !== -1) timeRow[nameIndex] = employeeName;
                  if (branchIndex !== -1) timeRow[branchIndex] = branch;
                  if (departmentIndex !== -1) timeRow[departmentIndex] = department;
                  employeeRows.push(timeRow);
                }
              }

              // Add employee rows to branch data
              branchData[branch].push(...employeeRows);

              // Skip the next 6 rows as we've already processed them
              i += 6;
            }
          }

          // Display branch information
          displayBranchResults();

          resolve();
        } catch (error) {
          reject(error);
        }
      };

      reader.onerror = function () {
        reject(new Error("Failed to read the file. Please try again."));
      };

      reader.readAsArrayBuffer(file);
    });
  }

  // Helper function to check if a value represents time data
  function isTimeValue(value) {
    if (!value) return true; // Empty cells are allowed in time rows
    if (value.toLowerCase() === "in" || value.toLowerCase() === "out" || 
        value.toLowerCase() === "hours" || value.toLowerCase() === "ot") {
      return true;
    }
    // Check for decimal numbers (Excel time format)
    if (!isNaN(value) && value.toString().includes('.')) {
      return true;
    }
    return false;
  }

  // Display branch results
  function displayBranchResults() {
    branchesList.innerHTML = "";

    // Sort branches alphabetically
    const sortedBranches = Object.keys(branchData).sort();

    sortedBranches.forEach((branch) => {
      const rowCount = branchData[branch].length - 1; // Subtract header row
      const employeeCount = Math.floor(rowCount / 7); // Each employee has 7 rows

      const branchItem = document.createElement("div");
      branchItem.className = "branch-item";

      const branchName = document.createElement("div");
      branchName.className = "branch-name";
      branchName.textContent = branch;

      const branchCount = document.createElement("div");
      branchCount.className = "branch-count";
      branchCount.textContent = `${employeeCount} employees (${rowCount} records)`;

      const downloadBtn = document.createElement("a");
      downloadBtn.className = "download-btn";
      downloadBtn.href = "#";
      downloadBtn.innerHTML = "<span>📥</span> Download";
      downloadBtn.addEventListener("click", function (e) {
        e.preventDefault();
        downloadBranchExcel(branch);
      });

      branchItem.appendChild(branchName);
      branchItem.appendChild(branchCount);
      branchItem.appendChild(downloadBtn);

      branchesList.appendChild(branchItem);
    });

    // Setup download all functionality
    downloadAllBtn.addEventListener("click", function (e) {
      e.preventDefault();
      downloadAllBranches();
    });
  }

  // Download Excel file for a specific branch with merged cells
  function downloadBranchExcel(branch) {
    try {
      // Create a new workbook
      const newWorkbook = XLSX.utils.book_new();
      
      // Convert branch data to worksheet
      const aoa = branchData[branch];
      const worksheet = XLSX.utils.aoa_to_sheet(aoa);
      
      // Process merged cells for the ID, Name, Branch, and Department columns
      const merges = [];
      
      // Find indices for relevant columns
      const headers = aoa[0];
      const idIndex = headers.findIndex(h => h && h.toString().toLowerCase() === "id");
      const nameIndex = headers.findIndex(h => h && h.toString().toLowerCase() === "name");
      const branchIndex = headers.findIndex(h => h && h.toString().toLowerCase() === "branch");
      const deptIndex = headers.findIndex(h => h && h.toString().toLowerCase() === "department");
      
      // Process employees (groups of 7 rows) for merging
      const columnIndices = [idIndex, nameIndex, branchIndex, deptIndex].filter(idx => idx !== -1);
      
      let rowIndex = 1; // Start after header row
      while (rowIndex < aoa.length) {
        // Process each group of 7 rows (1 employee record)
        const groupEndRow = Math.min(rowIndex + 6, aoa.length - 1);
        
        // For each column that needs merging
        columnIndices.forEach(colIndex => {
          if (colIndex >= 0) {
            // Add merge definition for this column across the 7 rows
            merges.push({
              s: { r: rowIndex, c: colIndex },
              e: { r: groupEndRow, c: colIndex }
            });
          }
        });
        
        // Move to next employee group
        rowIndex += 7;
      }
      
      // Add merges to worksheet
      worksheet['!merges'] = merges;

      // Add worksheet to workbook
      XLSX.utils.book_append_sheet(newWorkbook, worksheet, "Branch Data");

      // Generate Excel file
      const excelBuffer = XLSX.write(newWorkbook, {
        bookType: "xlsx",
        type: "array",
      });

      // Save file
      const blob = new Blob([excelBuffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      saveAs(blob, `Branch_${branch}.xlsx`);

      showAlert(
        `Successfully downloaded data for branch: ${branch}`,
        "success"
      );
    } catch (error) {
      showAlert(`Failed to download: ${error.message}`, "danger");
    }
  }

  // Download all branches as a zip file
  async function downloadAllBranches() {
    try {
      const zip = new JSZip();

      // Add each branch as a separate file
      Object.keys(branchData).forEach((branch) => {
        // Create a new workbook
        const newWorkbook = XLSX.utils.book_new();

        // Convert branch data to worksheet
        const aoa = branchData[branch];
        const worksheet = XLSX.utils.aoa_to_sheet(aoa);
        
        // Process merged cells for the ID, Name, Branch, and Department columns
        const merges = [];
        
        // Find indices for relevant columns
        const headers = aoa[0];
        const idIndex = headers.findIndex(h => h && h.toString().toLowerCase() === "id");
        const nameIndex = headers.findIndex(h => h && h.toString().toLowerCase() === "name");
        const branchIndex = headers.findIndex(h => h && h.toString().toLowerCase() === "branch");
        const deptIndex = headers.findIndex(h => h && h.toString().toLowerCase() === "department");
        
        // Process employees (groups of 7 rows) for merging
        const columnIndices = [idIndex, nameIndex, branchIndex, deptIndex].filter(idx => idx !== -1);
        
        let rowIndex = 1; // Start after header row
        while (rowIndex < aoa.length) {
          // Process each group of 7 rows (1 employee record)
          const groupEndRow = Math.min(rowIndex + 6, aoa.length - 1);
          
          // For each column that needs merging
          columnIndices.forEach(colIndex => {
            if (colIndex >= 0) {
              // Add merge definition for this column across the 7 rows
              merges.push({
                s: { r: rowIndex, c: colIndex },
                e: { r: groupEndRow, c: colIndex }
              });
            }
          });
          
          // Move to next employee group
          rowIndex += 7;
        }
        
        // Add merges to worksheet
        worksheet['!merges'] = merges;

        // Add worksheet to workbook
        XLSX.utils.book_append_sheet(newWorkbook, worksheet, "Branch Data");

        // Generate Excel file
        const excelBuffer = XLSX.write(newWorkbook, {
          bookType: "xlsx",
          type: "array",
        });

        // Add to zip
        zip.file(`Branch_${branch}.xlsx`, excelBuffer);
      });

      // Generate zip file
      const content = await zip.generateAsync({ type: "blob" });

      // Save zip file
      saveAs(content, "All_Branches.zip");

      showAlert(
        "Successfully downloaded all branches as a zip file",
        "success"
      );
    } catch (error) {
      showAlert(`Failed to download all branches: ${error.message}`, "danger");
    }
  }

  // Show alert message
  function showAlert(message, type) {
    const alert = document.createElement("div");
    alert.className = `alert alert-${type}`;
    alert.textContent = message;

    alertContainer.innerHTML = "";
    alertContainer.appendChild(alert);

    // Auto hide after 5 seconds
    setTimeout(() => {
      alert.style.opacity = "0";
      alert.style.transition = "opacity 0.5s";
      setTimeout(() => {
        if (alertContainer.contains(alert)) {
          alertContainer.removeChild(alert);
        }
      }, 500);
    }, 5000);
  }
});