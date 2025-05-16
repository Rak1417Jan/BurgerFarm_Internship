// let attendanceData = null;
// let employeesData = null;
// let additionalData = null;
// let processedData = null;
// let processedByDivision = {};

// // DOM elements
// const attendanceFileInput = document.getElementById("attendanceFile");
// const employeesFileInput = document.getElementById("employeesFile");
// const additionalFileInput = document.getElementById("additionalFile");
// const processBtn = document.getElementById("processBtn");
// const downloadAllBtn = document.getElementById("downloadAllBtn");
// const statusMessage = document.getElementById("statusMessage");
// const attendanceFileInfo = document.getElementById("attendanceFileInfo");
// const employeesFileInfo = document.getElementById("employeesFileInfo");
// const additionalFileInfo = document.getElementById("additionalFileInfo");
// const progressContainer = document.getElementById("progressContainer");
// const progressBar = document.getElementById("progressBar");
// const divisionDownloads = document.getElementById("divisionDownloads");
// const divisionButtons = document.getElementById("divisionButtons");

// // Event listeners for file inputs
// attendanceFileInput.addEventListener("change", (e) => {
//   if (e.target.files.length) {
//     attendanceFileInfo.textContent = `Selected: ${e.target.files[0].name}`;
//     attendanceFileInfo.classList.add("file-selected");
//     checkFilesReady();
//   } else {
//     attendanceFileInfo.textContent = "No file selected";
//     attendanceFileInfo.classList.remove("file-selected");
//   }
// });

// employeesFileInput.addEventListener("change", (e) => {
//   if (e.target.files.length) {
//     employeesFileInfo.textContent = `Selected: ${e.target.files[0].name}`;
//     employeesFileInfo.classList.add("file-selected");
//     checkFilesReady();
//   } else {
//     employeesFileInfo.textContent = "No file selected";
//     employeesFileInfo.classList.remove("file-selected");
//   }
// });

// additionalFileInput.addEventListener("change", (e) => {
//   if (e.target.files.length) {
//     additionalFileInfo.textContent = `Selected: ${e.target.files[0].name}`;
//     additionalFileInfo.classList.add("file-selected");
//     checkFilesReady();
//   } else {
//     additionalFileInfo.textContent = "No file selected";
//     additionalFileInfo.classList.remove("file-selected");
//   }
// });

// // Check if all files are ready for processing
// function checkFilesReady() {
//   if (
//     attendanceFileInput.files.length &&
//     employeesFileInput.files.length &&
//     additionalFileInput.files.length
//   ) {
//     processBtn.disabled = false;
//   } else {
//     processBtn.disabled = true;
//   }
//   // Reset download button when files change
//   downloadAllBtn.disabled = true;
//   processedData = null;
//   processedByDivision = {};
//   divisionDownloads.style.display = "none";
// }

// // Process button click handler
// processBtn.addEventListener("click", async () => {
//   try {
//     processBtn.disabled = true;
//     downloadAllBtn.disabled = true;
//     processBtn.textContent = "Processing...";
//     statusMessage.textContent = "Processing files, please wait...";
//     statusMessage.className = "status processing";

//     // Show progress indicator with immediate 100% progress
//     progressContainer.style.display = "block";
//     progressBar.style.width = "100%";

//     // Read all files
//     const [attendanceFile, employeesFile, additionalFile] = await Promise.all([
//       readFile(attendanceFileInput.files[0]),
//       readFile(employeesFileInput.files[0]),
//       readFile(additionalFileInput.files[0]),
//     ]);

//     // Parse the Excel files
//     const attendanceWorkbook = XLSX.read(attendanceFile, { type: "array" });
//     employeesData = parseExcel(employeesFile);

//     // Create additional data map that processes all sheets
//     additionalMap = createAdditionalMap(additionalFile);

//     // Log employee data headers for debugging
//     console.log("Employees file headers:", employeesData[0]);

//     // Process data for the first sheet (main processing)
//     const firstSheet = attendanceWorkbook.SheetNames[0];
//     const attendanceWorksheet = attendanceWorkbook.Sheets[firstSheet];
//     attendanceData = XLSX.utils.sheet_to_json(attendanceWorksheet, {
//       header: 1,
//       defval: "",
//     });

//     // Process the data
//     processedData = processAttendanceData(
//       attendanceData,
//       employeesData,
//       additionalMap
//     );

//     // Enable download button
//     downloadAllBtn.disabled = false;

//     // Show success message
//     statusMessage.textContent =
//       "Files processed successfully! Ready for download.";
//     statusMessage.className = "status success";

//     // Show division download buttons
//     showDivisionDownloads();
//   } catch (error) {
//     console.error("Error processing files:", error);
//     statusMessage.textContent = `Error: ${error.message}`;
//     statusMessage.className = "status error";
//     progressContainer.style.display = "none";
//   } finally {
//     processBtn.textContent = "Process Files";
//     processBtn.disabled = false;
//     progressBar.style.width = "0%";
//     progressContainer.style.display = "none";
//   }
// });

// // Download All button click handler
// downloadAllBtn.addEventListener("click", async () => {
//   if (!processedData || Object.keys(processedByDivision).length === 0) {
//     statusMessage.textContent =
//       "No processed data available. Please process files first.";
//     statusMessage.className = "status error";
//     return;
//   }

//   try {
//     statusMessage.textContent = "Generating ZIP file...";
//     statusMessage.className = "status processing";

//     const zip = new JSZip();
    
//     // Add the main processed file
//     const wb = XLSX.utils.book_new();
//     XLSX.utils.book_append_sheet(wb, processedData, "Processed Attendance");
//     const mainFileData = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
//     zip.file("Processed_Attendance_All.xlsx", mainFileData);

//     // Add division files
//     for (const [division, ws] of Object.entries(processedByDivision)) {
//       const divWb = XLSX.utils.book_new();
//       XLSX.utils.book_append_sheet(divWb, ws, "Processed Attendance");
//       const divFileData = XLSX.write(divWb, { bookType: 'xlsx', type: 'array' });
//       const safeDivisionName = division.replace(/[^a-zA-Z0-9]/g, '_');
//       zip.file(`Processed_Attendance_${safeDivisionName}.xlsx`, divFileData);
//     }

//     // Generate the ZIP file
//     const content = await zip.generateAsync({ type: 'blob' });
//     const url = URL.createObjectURL(content);
//     const a = document.createElement('a');
//     a.href = url;
//     a.download = 'Processed_Attendance_Files.zip';
//     document.body.appendChild(a);
//     a.click();
//     document.body.removeChild(a);
//     URL.revokeObjectURL(url);

//     statusMessage.textContent = "ZIP file downloaded successfully!";
//     statusMessage.className = "status success";
//   } catch (error) {
//     console.error("Error generating ZIP file:", error);
//     statusMessage.textContent = `Error generating ZIP file: ${error.message}`;
//     statusMessage.className = "status error";
//   }
// });

// // Show division download buttons
// function showDivisionDownloads() {
//   if (Object.keys(processedByDivision).length === 0) return;

//   divisionButtons.innerHTML = '';
  
//   // Create buttons for each division
//   for (const division of Object.keys(processedByDivision)) {
//     const btn = document.createElement('button');
//     btn.className = 'division-btn';
//     btn.textContent = `Download ${division}`;
//     btn.onclick = () => downloadDivisionFile(division);
//     divisionButtons.appendChild(btn);
//   }

//   divisionDownloads.style.display = 'block';
// }

// // Download a single division file
// function downloadDivisionFile(division) {
//   try {
//     if (!processedByDivision[division]) {
//       statusMessage.textContent = `No data available for division: ${division}`;
//       statusMessage.className = "status error";
//       return;
//     }

//     // Create workbook
//     const wb = XLSX.utils.book_new();
//     XLSX.utils.book_append_sheet(wb, processedByDivision[division], "Processed Attendance");

//     // Generate download
//     const safeDivisionName = division.replace(/[^a-zA-Z0-9]/g, '_');
//     XLSX.writeFile(wb, `Processed_Attendance_${safeDivisionName}.xlsx`);

//     statusMessage.textContent = `File for ${division} downloaded successfully!`;
//     statusMessage.className = "status success";
//   } catch (error) {
//     console.error(`Error downloading division ${division} file:`, error);
//     statusMessage.textContent = `Error downloading file: ${error.message}`;
//     statusMessage.className = "status error";
//   }
// }

// // Helper function to read a file as ArrayBuffer
// function readFile(file) {
//   return new Promise((resolve, reject) => {
//     const reader = new FileReader();
//     reader.onload = (e) => resolve(e.target.result);
//     reader.onerror = (e) => reject(new Error("Failed to read file"));
//     reader.readAsArrayBuffer(file);
//   });
// }

// // Helper function to parse Excel data from the first sheet
// function parseExcel(data) {
//   const workbook = XLSX.read(data, { type: "array" });
//   const firstSheetName = workbook.SheetNames[0];
//   const worksheet = workbook.Sheets[firstSheetName];
//   return XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
// }

// // Main processing function
// function processAttendanceData(attendance, employees, additionalMap) {
//   // Create a lookup map from employee ID to employee data
//   const employeeMap = createEmployeeMap(employees);

//   // Process each row in the attendance data
//   const processedRows = [];
//   const processedRowsByDivision = {};

//   // Process header row first
//   if (attendance.length === 0) {
//     throw new Error("Attendance file is empty");
//   }

//   // Find the index of "Department" and "Branch" columns in attendance data
//   const headerRow = attendance[0];
//   const deptIndex = headerRow.findIndex(
//     (col) => col && col.toString().trim().toLowerCase() === "department"
//   );
//   if (deptIndex === -1) {
//     throw new Error("Could not find 'Department' column in attendance file");
//   }

//   // Try to find Branch column if it exists
//   const branchIndex = headerRow.findIndex(
//     (col) =>
//       (col && col.toString().trim().toLowerCase() === "branch") ||
//       (col && col.toString().trim().toLowerCase().includes("branch"))
//   );

//   // Find the Days Payable column (second last column)
//   const daysPayableIndex = headerRow.length - 2;
//   if (daysPayableIndex < 0) {
//     throw new Error("Could not find 'Days Payable' column in attendance file");
//   }

//   // Create new header row
//   const newHeaderRow = [
//     ...headerRow.slice(0, deptIndex + 1), // Keep up to Department
//     "Designation",
//     "DOJ",
//     "Division",
//     ...headerRow.slice(-2), // Add last two columns
//     "Total Overtime", // Add the column from additional file
//     "Pending Offs OT" // Add the new column
//   ];

//   processedRows.push(newHeaderRow);

//   // Process data rows
//   for (let i = 1; i < attendance.length; i++) {
//     const row = attendance[i];
//     if (row.length === 0 || !row[0]) continue; // Skip empty rows

//     // Make sure we have valid data
//     const empId = row[0] ? row[0].toString().trim() : "";
//     if (!empId) continue; // Skip rows without valid employee ID

//     // Get branch if available
//     const branch =
//       branchIndex !== -1 && row[branchIndex]
//         ? row[branchIndex].toString().trim()
//         : "";

//     const employeeInfo = employeeMap[empId] || {};

//     // Look up additional info using employee ID
//     let overtimeValue = "";
//     if (additionalMap[empId]) {
//       overtimeValue = additionalMap[empId];
//       console.log(
//         `Found overtime value for employee ${empId}: ${overtimeValue}`
//       );
//     } else {
//       console.log(`No overtime value found for employee ${empId}`);
//     }

//     // Calculate Pending Offs OT based on Days Payable
//     const daysPayable = parseFloat(row[daysPayableIndex]) || 0;
//     let pendingOffsOT = 0;
    
//     if (daysPayable <= 6) {
//       pendingOffsOT = 0;
//     } else if (daysPayable <= 13) {
//       pendingOffsOT = 9;
//     } else if (daysPayable <= 20) {
//       pendingOffsOT = 18;
//     } else if (daysPayable <= 23) {
//       pendingOffsOT = 27;
//     } else {
//       pendingOffsOT = 36;
//     }

//     // Create new row with selected columns
//     const newRow = [
//       ...row.slice(0, deptIndex + 1), // Keep up to Department
//       employeeInfo.Designation || "", // Add Designation
//       employeeInfo.DOJ || "", // Add DOJ
//       employeeInfo.Division || "", // Add Division
//       ...row.slice(-2), // Add last two columns
//       overtimeValue, // Add Total Overtime data
//       pendingOffsOT // Add Pending Offs OT
//     ];

//     processedRows.push(newRow);

//     // Add to division-specific data
//     const division = employeeInfo.Division || "Unknown";
//     if (!processedRowsByDivision[division]) {
//       // Create header row for division
//       processedRowsByDivision[division] = [newHeaderRow];
//     }
//     processedRowsByDivision[division].push(newRow);
//   }

//   // Convert all data to worksheet format
//   const ws = XLSX.utils.aoa_to_sheet(processedRows);
  
//   // Convert division data to worksheets
//   for (const [division, rows] of Object.entries(processedRowsByDivision)) {
//     processedByDivision[division] = XLSX.utils.aoa_to_sheet(rows);
//   }

//   return ws;
// }

// // Create a map of employee data from the employees file
// function createEmployeeMap(employees) {
//   const map = {};

//   if (employees.length < 2) {
//     throw new Error("Employees file doesn't contain enough data");
//   }

//   const headers = employees[0].map((h) =>
//     h && h.toString ? h.toString().trim() : ""
//   );

//   // Find column indices
//   const empIdIndex = headers.findIndex(
//     (h) =>
//       h.toLowerCase().includes("employee id") ||
//       h.toLowerCase().includes("emp id") ||
//       h.toLowerCase().includes("emp code") ||
//       h.toLowerCase() === "empcode" ||
//       h.toLowerCase() === "code" ||
//       h.toLowerCase() === "employee code" ||
//       h.toLowerCase() === "employee i'd" ||
//       h.toLowerCase() === "employ id"
//   );
//   const designationIndex = headers.findIndex((h) =>
//     h.toLowerCase().includes("designation")
//   );
//   const dojIndex = headers.findIndex(
//     (h) => h.toLowerCase() === "doj" || h.toLowerCase().includes("date of join")
//   );
//   const divisionIndex = headers.findIndex((h) =>
//     h.toLowerCase().includes("division")
//   );

//   if (empIdIndex === -1) {
//     throw new Error("Employee ID column not found in employees file");
//   }

//   if (designationIndex === -1 || dojIndex === -1 || divisionIndex === -1) {
//     throw new Error(
//       "Required columns (Designation, DOJ, or Division) not found in employees file"
//     );
//   }

//   // Process each employee row
//   for (let i = 1; i < employees.length; i++) {
//     const row = employees[i];
//     if (row.length <= empIdIndex || !row[empIdIndex]) continue;

//     const empId = row[empIdIndex].toString().trim();

//     map[empId] = {
//       Designation: row[designationIndex] || "",
//       DOJ: formatDate(row[dojIndex]) || "",
//       Division: row[divisionIndex] || "",
//     };
//   }

//   return map;
// }

// // Create a map of additional data (Total Overtime column) using employee ID as keys
// function createAdditionalMap(additionalFile) {
//   const map = {};

//   try {
//     // Read the additional file as a workbook to access multiple sheets
//     const additionalWorkbook = XLSX.read(additionalFile, { type: "array" });
//     console.log("Additional file sheets:", additionalWorkbook.SheetNames);

//     // Process each sheet in the additional file
//     additionalWorkbook.SheetNames.forEach((sheetName) => {
//       console.log(`Processing sheet: ${sheetName}`);

//       // Extract data from this sheet
//       const worksheet = additionalWorkbook.Sheets[sheetName];
//       const sheetData = XLSX.utils.sheet_to_json(worksheet, {
//         header: 1,
//         defval: "",
//       });

//       if (sheetData.length < 2) {
//         console.warn(`Sheet ${sheetName} has insufficient data, skipping`);
//         return; // Skip this sheet
//       }

//       // Get headers
//       const headers = sheetData[0].map((h) =>
//         h && h.toString ? h.toString().trim() : ""
//       );
//       console.log(`Sheet ${sheetName} headers:`, headers);

//       // Find employee ID column (look for variations)
//       const empIdIndex = headers.findIndex(
//         (h) =>
//           h.toLowerCase().includes("employee id") ||
//           h.toLowerCase().includes("emp id") ||
//           h.toLowerCase().includes("emp code") ||
//           h.toLowerCase() === "empcode" ||
//           h.toLowerCase() === "code" ||
//           h.toLowerCase() === "employee code" ||
//           h.toLowerCase() === "employee i'd" ||
//           h.toLowerCase() === "employ id"
//       );

//       // If no ID column found, try first column
//       const idColIndex = empIdIndex !== -1 ? empIdIndex : 0;

//       // Find the "Total Overtime" column index
//       const overtimeIndex = headers.findIndex(
//         (h) => h && h.toString().trim().toLowerCase() === "total overtime"
//       );

//       // Skip this sheet if we can't find the Total Overtime column
//       if (overtimeIndex === -1) {
//         console.warn(
//           `Sheet ${sheetName} doesn't contain 'Total Overtime' column, skipping`
//         );
//         return;
//       }

//       // Process each row in this sheet
//       for (let i = 1; i < sheetData.length; i++) {
//         const row = sheetData[i];
//         if (row.length <= idColIndex || !row[idColIndex]) continue;

//         const empId = row[idColIndex].toString().trim();

//         // Get the overtime value and normalize it
//         let overtimeValue = row[overtimeIndex] || "";

//         // Convert the overtime value to a properly formatted number
//         overtimeValue = normalizeNumberValue(overtimeValue);

//         // Store the overtime value keyed by employee ID
//         map[empId] = overtimeValue;

//         // Debug
//         console.log(
//           `Added employee ${empId} from branch ${sheetName} with overtime value ${overtimeValue}`
//         );
//       }
//     });

//     return map;
//   } catch (error) {
//     console.error("Error processing additional file:", error);
//     throw new Error(`Failed to process additional file: ${error.message}`);
//   }
// }

// // Helper function to normalize number values and prevent scientific notation
// function normalizeNumberValue(value) {
//   // If empty, return empty string
//   if (value === "" || value === null || value === undefined) {
//     return "";
//   }

//   try {
//     // If it's already a number, format it
//     if (typeof value === "number") {
//       // Convert to fixed decimal to avoid scientific notation
//       return Number(value).toFixed(2);
//     }

//     // If it's a string that might be a number
//     if (typeof value === "string") {
//       // Try to convert to a number
//       const num = parseFloat(value.replace(/,/g, ""));
//       if (!isNaN(num)) {
//         // Format to 2 decimal places
//         return num.toFixed(2);
//       }
//       // If not a valid number, return original string
//       return value;
//     }

//     // For any other type, return as string
//     return String(value);
//   } catch (error) {
//     console.warn(`Error normalizing value: ${value}`, error);
//     return String(value);
//   }
// }

// // Helper function to format date
// function formatDate(dateStr) {
//   if (!dateStr) return "";

//   // Handle Excel serial date numbers (DATE format in Excel)
//   if (typeof dateStr === "number") {
//     // Excel date is number of days since 1900-01-01 (with 1900 incorrectly treated as leap year)
//     const excelEpoch = new Date(1899, 11, 31);
//     const date = new Date(excelEpoch.getTime() + dateStr * 24 * 60 * 60 * 1000);

//     // Format as DD-MM-YYYY
//     const day = String(date.getDate()).padStart(2, "0");
//     const month = String(date.getMonth() + 1).padStart(2, "0");
//     const year = date.getFullYear();
//     return `${day}-${month}-${year}`;
//   }

//   // Handle string dates that might already be in DD-MM-YYYY format
//   if (typeof dateStr === "string") {
//     // Check if it's already in DD-MM-YYYY format
//     const ddMmYyyyFormat = /^(\d{2})-(\d{2})-(\d{4})$/;
//     const match = dateStr.match(ddMmYyyyFormat);
//     if (match) {
//       return dateStr; // Return as-is if already in correct format
//     }

//     // Try to parse other common date formats and convert to DD-MM-YYYY
//     const dateFormats = [
//       /(\d{4})-(\d{2})-(\d{2})/, // YYYY-MM-DD
//       /(\d{2})\/(\d{2})\/(\d{4})/, // MM/DD/YYYY or DD/MM/YYYY
//       /(\d{4})\/(\d{2})\/(\d{2})/, // YYYY/MM/DD
//     ];

//     for (const format of dateFormats) {
//       const match = dateStr.match(format);
//       if (match) {
//         const parts = match.slice(1).map((part) => part.padStart(2, "0"));
//         // Determine the order of the parts based on the format
//         let day, month, year;
//         if (format.toString().includes("YYYY-MM-DD")) {
//           [year, month, day] = parts;
//         } else if (format.toString().includes("YYYY/MM/DD")) {
//           [year, month, day] = parts;
//         } else {
//           // Assume DD/MM/YYYY format for simplicity
//           [day, month, year] = parts;
//         }
//         return `${day}-${month}-${year}`;
//       }
//     }

//     // If no format matched, return as-is
//     return dateStr;
//   }

//   return "";
// }

let attendanceData = null;
let employeesData = null;
let additionalData = null;
let processedData = null;
let processedByDivision = {};

// DOM elements
const attendanceFileInput = document.getElementById("attendanceFile");
const employeesFileInput = document.getElementById("employeesFile");
const additionalFileInput = document.getElementById("additionalFile");
const processBtn = document.getElementById("processBtn");
const downloadAllBtn = document.getElementById("downloadAllBtn");
const statusMessage = document.getElementById("statusMessage");
const attendanceFileInfo = document.getElementById("attendanceFileInfo");
const employeesFileInfo = document.getElementById("employeesFileInfo");
const additionalFileInfo = document.getElementById("additionalFileInfo");
const progressContainer = document.getElementById("progressContainer");
const progressBar = document.getElementById("progressBar");
const divisionDownloads = document.getElementById("divisionDownloads");
const divisionButtons = document.getElementById("divisionButtons");

// Event listeners for file inputs
attendanceFileInput.addEventListener("change", (e) => {
  if (e.target.files.length) {
    attendanceFileInfo.textContent = `Selected: ${e.target.files[0].name}`;
    attendanceFileInfo.classList.add("file-selected");
    checkFilesReady();
  } else {
    attendanceFileInfo.textContent = "No file selected";
    attendanceFileInfo.classList.remove("file-selected");
  }
});

employeesFileInput.addEventListener("change", (e) => {
  if (e.target.files.length) {
    employeesFileInfo.textContent = `Selected: ${e.target.files[0].name}`;
    employeesFileInfo.classList.add("file-selected");
    checkFilesReady();
  } else {
    employeesFileInfo.textContent = "No file selected";
    employeesFileInfo.classList.remove("file-selected");
  }
});

additionalFileInput.addEventListener("change", (e) => {
  if (e.target.files.length) {
    additionalFileInfo.textContent = `Selected: ${e.target.files[0].name}`;
    additionalFileInfo.classList.add("file-selected");
    checkFilesReady();
  } else {
    additionalFileInfo.textContent = "No file selected";
    additionalFileInfo.classList.remove("file-selected");
  }
});

// Check if all files are ready for processing
function checkFilesReady() {
  if (
    attendanceFileInput.files.length &&
    employeesFileInput.files.length &&
    additionalFileInput.files.length
  ) {
    processBtn.disabled = false;
  } else {
    processBtn.disabled = true;
  }
  // Reset download button when files change
  downloadAllBtn.disabled = true;
  processedData = null;
  processedByDivision = {};
  divisionDownloads.style.display = "none";
}

// Process button click handler
processBtn.addEventListener("click", async () => {
  try {
    processBtn.disabled = true;
    downloadAllBtn.disabled = true;
    processBtn.textContent = "Processing...";
    statusMessage.textContent = "Processing files, please wait...";
    statusMessage.className = "status processing";

    // Show progress indicator with immediate 100% progress
    progressContainer.style.display = "block";
    progressBar.style.width = "100%";

    // Read all files
    const [attendanceFile, employeesFile, additionalFile] = await Promise.all([
      readFile(attendanceFileInput.files[0]),
      readFile(employeesFileInput.files[0]),
      readFile(additionalFileInput.files[0]),
    ]);

    // Parse the Excel files
    const attendanceWorkbook = XLSX.read(attendanceFile, { type: "array" });
    employeesData = parseExcel(employeesFile);

    // Create additional data map that processes all sheets
    additionalMap = createAdditionalMap(additionalFile);

    // Log employee data headers for debugging
    console.log("Employees file headers:", employeesData[0]);

    // Process data for the first sheet (main processing)
    const firstSheet = attendanceWorkbook.SheetNames[0];
    const attendanceWorksheet = attendanceWorkbook.Sheets[firstSheet];
    attendanceData = XLSX.utils.sheet_to_json(attendanceWorksheet, {
      header: 1,
      defval: "",
    });

    // Process the data
    processedData = processAttendanceData(
      attendanceData,
      employeesData,
      additionalMap
    );

    // Enable download button
    downloadAllBtn.disabled = false;

    // Show success message
    statusMessage.textContent =
      "Files processed successfully! Ready for download.";
    statusMessage.className = "status success";

    // Show division download buttons
    showDivisionDownloads();
  } catch (error) {
    console.error("Error processing files:", error);
    statusMessage.textContent = `Error: ${error.message}`;
    statusMessage.className = "status error";
    progressContainer.style.display = "none";
  } finally {
    processBtn.textContent = "Process Files";
    processBtn.disabled = false;
    progressBar.style.width = "0%";
    progressContainer.style.display = "none";
  }
});

// Download All button click handler
downloadAllBtn.addEventListener("click", async () => {
  if (!processedData || Object.keys(processedByDivision).length === 0) {
    statusMessage.textContent =
      "No processed data available. Please process files first.";
    statusMessage.className = "status error";
    return;
  }

  try {
    statusMessage.textContent = "Generating ZIP file...";
    statusMessage.className = "status processing";

    const zip = new JSZip();
    
    // Add the main processed file
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, processedData, "Processed Attendance");
    const mainFileData = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    zip.file("Processed_Attendance_All.xlsx", mainFileData);

    // Add division files
    for (const [division, ws] of Object.entries(processedByDivision)) {
      const divWb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(divWb, ws, "Processed Attendance");
      const divFileData = XLSX.write(divWb, { bookType: 'xlsx', type: 'array' });
      const safeDivisionName = division.replace(/[^a-zA-Z0-9]/g, '_');
      zip.file(`Processed_Attendance_${safeDivisionName}.xlsx`, divFileData);
    }

    // Generate the ZIP file
    const content = await zip.generateAsync({ type: 'blob' });
    const url = URL.createObjectURL(content);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'Processed_Attendance_Files.zip';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    statusMessage.textContent = "ZIP file downloaded successfully!";
    statusMessage.className = "status success";
  } catch (error) {
    console.error("Error generating ZIP file:", error);
    statusMessage.textContent = `Error generating ZIP file: ${error.message}`;
    statusMessage.className = "status error";
  }
});

// Show division download buttons
function showDivisionDownloads() {
  if (Object.keys(processedByDivision).length === 0) return;

  divisionButtons.innerHTML = '';
  
  // Create buttons for each division
  for (const division of Object.keys(processedByDivision)) {
    const btn = document.createElement('button');
    btn.className = 'division-btn';
    btn.textContent = `Download ${division}`;
    btn.onclick = () => downloadDivisionFile(division);
    divisionButtons.appendChild(btn);
  }

  divisionDownloads.style.display = 'block';
}

// Download a single division file
function downloadDivisionFile(division) {
  try {
    if (!processedByDivision[division]) {
      statusMessage.textContent = `No data available for division: ${division}`;
      statusMessage.className = "status error";
      return;
    }

    // Create workbook
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, processedByDivision[division], "Processed Attendance");

    // Generate download
    const safeDivisionName = division.replace(/[^a-zA-Z0-9]/g, '_');
    XLSX.writeFile(wb, `Processed_Attendance_${safeDivisionName}.xlsx`);

    statusMessage.textContent = `File for ${division} downloaded successfully!`;
    statusMessage.className = "status success";
  } catch (error) {
    console.error(`Error downloading division ${division} file:`, error);
    statusMessage.textContent = `Error downloading file: ${error.message}`;
    statusMessage.className = "status error";
  }
}

// Helper function to read a file as ArrayBuffer
function readFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => resolve(e.target.result);
    reader.onerror = (e) => reject(new Error("Failed to read file"));
    reader.readAsArrayBuffer(file);
  });
}

// Helper function to parse Excel data from the first sheet
function parseExcel(data) {
  const workbook = XLSX.read(data, { type: "array" });
  const firstSheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[firstSheetName];
  return XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
}

// Main processing function
function processAttendanceData(attendance, employees, additionalMap) {
  // Create a lookup map from employee ID to employee data
  const employeeMap = createEmployeeMap(employees);

  // Process each row in the attendance data
  const processedRows = [];
  const processedRowsByDivision = {};

  // Process header row first
  if (attendance.length === 0) {
    throw new Error("Attendance file is empty");
  }

  // Find the index of "Department" and "Branch" columns in attendance data
  const headerRow = attendance[0];
  const deptIndex = headerRow.findIndex(
    (col) => col && col.toString().trim().toLowerCase() === "department"
  );
  if (deptIndex === -1) {
    throw new Error("Could not find 'Department' column in attendance file");
  }

  // Try to find Branch column if it exists
  const branchIndex = headerRow.findIndex(
    (col) =>
      (col && col.toString().trim().toLowerCase() === "branch") ||
      (col && col.toString().trim().toLowerCase().includes("branch"))
  );

  // Find the Days Payable column (second last column)
  const daysPayableIndex = headerRow.length - 2;
  if (daysPayableIndex < 0) {
    throw new Error("Could not find 'Days Payable' column in attendance file");
  }

  // Create new header row
  const newHeaderRow = [
    ...headerRow.slice(0, deptIndex + 1), // Keep up to Department
    "Designation",
    "DOJ",
    "Division",
    ...headerRow.slice(-2), // Add last two columns
    "Total Overtime", // Add the column from additional file
    "Pending Offs OT" // Add the new column
  ];

  processedRows.push(newHeaderRow);

  // Process data rows
  for (let i = 1; i < attendance.length; i++) {
    const row = attendance[i];
    if (row.length === 0 || !row[0]) continue; // Skip empty rows

    // Make sure we have valid data
    const empId = row[0] ? row[0].toString().trim() : "";
    if (!empId) continue; // Skip rows without valid employee ID

    // Get branch if available
    const branch =
      branchIndex !== -1 && row[branchIndex]
        ? row[branchIndex].toString().trim()
        : "";

    const employeeInfo = employeeMap[empId] || {};

    // Look up additional info using employee ID
    let overtimeValue = "";
    if (additionalMap[empId]) {
      overtimeValue = additionalMap[empId];
      console.log(
        `Found overtime value for employee ${empId}: ${overtimeValue}`
      );
    } else {
      console.log(`No overtime value found for employee ${empId}`);
    }

    // Extract the number from Days Payable (format: "Present:10.5")
    let daysPayableStr = row[daysPayableIndex] ? row[daysPayableIndex].toString().trim() : "";
    let daysPayable = 0;
    
    if (daysPayableStr.includes(":")) {
      const parts = daysPayableStr.split(":");
      if (parts.length > 1) {
        daysPayable = parseFloat(parts[1]) || 0;
      }
    } else {
      daysPayable = parseFloat(daysPayableStr) || 0;
    }

    // Calculate Pending Offs OT based on Days Payable
    let pendingOffsOT = 0;
    
    if (daysPayable <= 6) {
      pendingOffsOT = 0;
    } else if (daysPayable <= 13) {
      pendingOffsOT = 9;
    } else if (daysPayable <= 20) {
      pendingOffsOT = 18;
    } else if (daysPayable <= 23) {
      pendingOffsOT = 27;
    } else {
      pendingOffsOT = 36;
    }

    // Create new row with selected columns
    const newRow = [
      ...row.slice(0, deptIndex + 1), // Keep up to Department
      employeeInfo.Designation || "", // Add Designation
      employeeInfo.DOJ || "", // Add DOJ
      employeeInfo.Division || "", // Add Division
      ...row.slice(-2), // Add last two columns
      overtimeValue, // Add Total Overtime data
      pendingOffsOT // Add Pending Offs OT
    ];

    processedRows.push(newRow);

    // Add to division-specific data
    const division = employeeInfo.Division || "Unknown";
    if (!processedRowsByDivision[division]) {
      // Create header row for division
      processedRowsByDivision[division] = [newHeaderRow];
    }
    processedRowsByDivision[division].push(newRow);
  }

  // Convert all data to worksheet format
  const ws = XLSX.utils.aoa_to_sheet(processedRows);
  
  // Convert division data to worksheets
  for (const [division, rows] of Object.entries(processedRowsByDivision)) {
    processedByDivision[division] = XLSX.utils.aoa_to_sheet(rows);
  }

  return ws;
}

// Create a map of employee data from the employees file
function createEmployeeMap(employees) {
  const map = {};

  if (employees.length < 2) {
    throw new Error("Employees file doesn't contain enough data");
  }

  const headers = employees[0].map((h) =>
    h && h.toString ? h.toString().trim() : ""
  );

  // Find column indices
  const empIdIndex = headers.findIndex(
    (h) =>
      h.toLowerCase().includes("employee id") ||
      h.toLowerCase().includes("emp id") ||
      h.toLowerCase().includes("emp code") ||
      h.toLowerCase() === "empcode" ||
      h.toLowerCase() === "code" ||
      h.toLowerCase() === "employee code" ||
      h.toLowerCase() === "employee i'd" ||
      h.toLowerCase() === "employ id"
  );
  const designationIndex = headers.findIndex((h) =>
    h.toLowerCase().includes("designation")
  );
  const dojIndex = headers.findIndex(
    (h) => h.toLowerCase() === "doj" || h.toLowerCase().includes("date of join")
  );
  const divisionIndex = headers.findIndex((h) =>
    h.toLowerCase().includes("division")
  );

  if (empIdIndex === -1) {
    throw new Error("Employee ID column not found in employees file");
  }

  if (designationIndex === -1 || dojIndex === -1 || divisionIndex === -1) {
    throw new Error(
      "Required columns (Designation, DOJ, or Division) not found in employees file"
    );
  }

  // Process each employee row
  for (let i = 1; i < employees.length; i++) {
    const row = employees[i];
    if (row.length <= empIdIndex || !row[empIdIndex]) continue;

    const empId = row[empIdIndex].toString().trim();

    map[empId] = {
      Designation: row[designationIndex] || "",
      DOJ: formatDate(row[dojIndex]) || "",
      Division: row[divisionIndex] || "",
    };
  }

  return map;
}

// Create a map of additional data (Total Overtime column) using employee ID as keys
function createAdditionalMap(additionalFile) {
  const map = {};

  try {
    // Read the additional file as a workbook to access multiple sheets
    const additionalWorkbook = XLSX.read(additionalFile, { type: "array" });
    console.log("Additional file sheets:", additionalWorkbook.SheetNames);

    // Process each sheet in the additional file
    additionalWorkbook.SheetNames.forEach((sheetName) => {
      console.log(`Processing sheet: ${sheetName}`);

      // Extract data from this sheet
      const worksheet = additionalWorkbook.Sheets[sheetName];
      const sheetData = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
        defval: "",
      });

      if (sheetData.length < 2) {
        console.warn(`Sheet ${sheetName} has insufficient data, skipping`);
        return; // Skip this sheet
      }

      // Get headers
      const headers = sheetData[0].map((h) =>
        h && h.toString ? h.toString().trim() : ""
      );
      console.log(`Sheet ${sheetName} headers:`, headers);

      // Find employee ID column (look for variations)
      const empIdIndex = headers.findIndex(
        (h) =>
          h.toLowerCase().includes("employee id") ||
          h.toLowerCase().includes("emp id") ||
          h.toLowerCase().includes("emp code") ||
          h.toLowerCase() === "empcode" ||
          h.toLowerCase() === "code" ||
          h.toLowerCase() === "employee code" ||
          h.toLowerCase() === "employee i'd" ||
          h.toLowerCase() === "employ id"
      );

      // If no ID column found, try first column
      const idColIndex = empIdIndex !== -1 ? empIdIndex : 0;

      // Find the "Total Overtime" column index
      const overtimeIndex = headers.findIndex(
        (h) => h && h.toString().trim().toLowerCase() === "total overtime"
      );

      // Skip this sheet if we can't find the Total Overtime column
      if (overtimeIndex === -1) {
        console.warn(
          `Sheet ${sheetName} doesn't contain 'Total Overtime' column, skipping`
        );
        return;
      }

      // Process each row in this sheet
      for (let i = 1; i < sheetData.length; i++) {
        const row = sheetData[i];
        if (row.length <= idColIndex || !row[idColIndex]) continue;

        const empId = row[idColIndex].toString().trim();

        // Get the overtime value and normalize it
        let overtimeValue = row[overtimeIndex] || "";

        // Convert the overtime value to a properly formatted number
        overtimeValue = normalizeNumberValue(overtimeValue);

        // Store the overtime value keyed by employee ID
        map[empId] = overtimeValue;

        // Debug
        console.log(
          `Added employee ${empId} from branch ${sheetName} with overtime value ${overtimeValue}`
        );
      }
    });

    return map;
  } catch (error) {
    console.error("Error processing additional file:", error);
    throw new Error(`Failed to process additional file: ${error.message}`);
  }
}

// Helper function to normalize number values and prevent scientific notation
function normalizeNumberValue(value) {
  // If empty, return empty string
  if (value === "" || value === null || value === undefined) {
    return "";
  }

  try {
    // If it's already a number, format it
    if (typeof value === "number") {
      // Convert to fixed decimal to avoid scientific notation
      return Number(value).toFixed(2);
    }

    // If it's a string that might be a number
    if (typeof value === "string") {
      // Try to convert to a number
      const num = parseFloat(value.replace(/,/g, ""));
      if (!isNaN(num)) {
        // Format to 2 decimal places
        return num.toFixed(2);
      }
      // If not a valid number, return original string
      return value;
    }

    // For any other type, return as string
    return String(value);
  } catch (error) {
    console.warn(`Error normalizing value: ${value}`, error);
    return String(value);
  }
}

// Helper function to format date
function formatDate(dateStr) {
  if (!dateStr) return "";

  // Handle Excel serial date numbers (DATE format in Excel)
  if (typeof dateStr === "number") {
    // Excel date is number of days since 1900-01-01 (with 1900 incorrectly treated as leap year)
    const excelEpoch = new Date(1899, 11, 31);
    const date = new Date(excelEpoch.getTime() + dateStr * 24 * 60 * 60 * 1000);

    // Format as DD-MM-YYYY
    const day = String(date.getDate()).padStart(2, "0");
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const year = date.getFullYear();
    return `${day}-${month}-${year}`;
  }

  // Handle string dates that might already be in DD-MM-YYYY format
  if (typeof dateStr === "string") {
    // Check if it's already in DD-MM-YYYY format
    const ddMmYyyyFormat = /^(\d{2})-(\d{2})-(\d{4})$/;
    const match = dateStr.match(ddMmYyyyFormat);
    if (match) {
      return dateStr; // Return as-is if already in correct format
    }

    // Try to parse other common date formats and convert to DD-MM-YYYY
    const dateFormats = [
      /(\d{4})-(\d{2})-(\d{2})/, // YYYY-MM-DD
      /(\d{2})\/(\d{2})\/(\d{4})/, // MM/DD/YYYY or DD/MM/YYYY
      /(\d{4})\/(\d{2})\/(\d{2})/, // YYYY/MM/DD
    ];

    for (const format of dateFormats) {
      const match = dateStr.match(format);
      if (match) {
        const parts = match.slice(1).map((part) => part.padStart(2, "0"));
        // Determine the order of the parts based on the format
        let day, month, year;
        if (format.toString().includes("YYYY-MM-DD")) {
          [year, month, day] = parts;
        } else if (format.toString().includes("YYYY/MM/DD")) {
          [year, month, day] = parts;
        } else {
          // Assume DD/MM/YYYY format for simplicity
          [day, month, year] = parts;
        }
        return `${day}-${month}-${year}`;
      }
    }

    // If no format matched, return as-is
    return dateStr;
  }

  return "";
}