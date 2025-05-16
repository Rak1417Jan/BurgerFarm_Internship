// // Global variables
// let convertedData = [];
// let fileName = "";
// let employeeIdMap = {};
// let branchEmployeeIds = {};
// let branchDetailsMap = {};

// // Handle onboarding file input change
// document
//   .getElementById("onboardingFileInput")
//   .addEventListener("change", function (e) {
//     const file = e.target.files[0];
//     if (!file) return;

//     document.getElementById(
//       "onboardingFileName"
//     ).textContent = `Selected: ${file.name}`;
//     document.getElementById("statusMessage").textContent =
//       "Please upload all files to proceed";
//     document.getElementById("statusMessage").className = "status";
//   });

// // Handle employee ID file input change
// document
//   .getElementById("employeeIdFileInput")
//   .addEventListener("change", function (e) {
//     const file = e.target.files[0];
//     if (!file) return;

//     document.getElementById(
//       "employeeIdFileName"
//     ).textContent = `Selected: ${file.name}`;
//     document.getElementById("statusMessage").textContent =
//       "Processing employee ID file...";
//     document.getElementById("statusMessage").className = "status";

//     const reader = new FileReader();
//     reader.onload = function (e) {
//       try {
//         const data = new Uint8Array(e.target.result);
//         const workbook = XLSX.read(data, { type: "array" });

//         // Assuming the first sheet is the one we want
//         const firstSheetName = workbook.SheetNames[0];
//         const worksheet = workbook.Sheets[firstSheetName];

//         // Convert to JSON
//         const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

//         // Process the employee ID data
//         processEmployeeIdData(jsonData);

//         document.getElementById("statusMessage").textContent =
//           "Employee ID file processed. Please upload other files if not done yet.";
//         document.getElementById("statusMessage").className = "status";

//         // Check if all files are uploaded to process
//         if (areAllFilesUploaded()) {
//           processAllFiles();
//         }
//       } catch (error) {
//         document.getElementById("statusMessage").textContent =
//           "Error processing employee ID file: " + error.message;
//         document.getElementById("statusMessage").className = "status error";
//         console.error(error);
//       }
//     };
//     reader.readAsArrayBuffer(file);
//   });

// // Handle branch details file input change
// document
//   .getElementById("branchDetailsFileInput")
//   .addEventListener("change", function (e) {
//     const file = e.target.files[0];
//     if (!file) return;

//     document.getElementById(
//       "branchDetailsFileName"
//     ).textContent = `Selected: ${file.name}`;
//     document.getElementById("statusMessage").textContent =
//       "Processing branch details file...";
//     document.getElementById("statusMessage").className = "status";

//     const reader = new FileReader();
//     reader.onload = function (e) {
//       try {
//         const data = new Uint8Array(e.target.result);
//         const workbook = XLSX.read(data, { type: "array" });

//         // Assuming the first sheet is the one we want
//         const firstSheetName = workbook.SheetNames[0];
//         const worksheet = workbook.Sheets[firstSheetName];

//         // Convert to JSON
//         const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

//         // Process the branch details data
//         processBranchDetailsData(jsonData);

//         document.getElementById("statusMessage").textContent =
//           "Branch details file processed. Please upload other files if not done yet.";
//         document.getElementById("statusMessage").className = "status";

//         // Check if all files are uploaded to process
//         if (areAllFilesUploaded()) {
//           processAllFiles();
//         }
//       } catch (error) {
//         document.getElementById("statusMessage").textContent =
//           "Error processing branch details file: " + error.message;
//         document.getElementById("statusMessage").className = "status error";
//         console.error(error);
//       }
//     };
//     reader.readAsArrayBuffer(file);
//   });

// // Process branch details data and build the mapping
// function processBranchDetailsData(data) {
//   branchDetailsMap = {};

//   if (data.length < 2) return;

//   // Get headers
//   const headers = data[0].map((h) =>
//     typeof h === "string" ? h.toUpperCase() : h
//   );
//   const branchIndex = headers.indexOf("BRANCH");
//   const locationIndex = headers.indexOf("LOCATION");
//   const divisionIndex = headers.indexOf("DIVISION");

//   // Process each row
//   for (let i = 1; i < data.length; i++) {
//     const row = data[i];
//     if (row.length <= Math.max(branchIndex, locationIndex, divisionIndex)) continue;

//     const branch = formatCellValue(row[branchIndex]);
//     const location = formatCellValue(row[locationIndex]);
//     const division = formatCellValue(row[divisionIndex]);

//     // Store mapping of branch to location and division
//     if (branch) {
//       branchDetailsMap[branch] = {
//         location: location || "",
//         division: division || ""
//       };
//     }
//   }
// }

// // Check if all required files are uploaded
// function areAllFilesUploaded() {
//   return (
//     document.getElementById("onboardingFileInput").files.length > 0 &&
//     document.getElementById("employeeIdFileInput").files.length > 0 &&
//     document.getElementById("branchDetailsFileInput").files.length > 0
//   );
// }

// // Process all files when all are uploaded
// function processAllFiles() {
//   const onboardingFile = document.getElementById("onboardingFileInput").files[0];
//   if (!onboardingFile) return;

//   fileName = onboardingFile.name.replace(/\.[^/.]+$/, "") + "_converted.xlsx";
//   document.getElementById("statusMessage").textContent = "Processing all files...";
//   document.getElementById("statusMessage").className = "status";

//   const reader = new FileReader();
//   reader.onload = function (e) {
//     try {
//       const data = new Uint8Array(e.target.result);
//       const workbook = XLSX.read(data, { type: "array" });

//       // Assuming the first sheet is the one we want
//       const firstSheetName = workbook.SheetNames[0];
//       const worksheet = workbook.Sheets[firstSheetName];

//       // Convert to JSON
//       const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

//       // Process the data (now returns an object with both sheets)
//       const convertedData = convertOnboardingData(jsonData);

//       // Display preview (showing just the first sheet)
//       displayPreview(convertedData.employeeData);

//       document.getElementById("statusMessage").textContent = "Files processed successfully!";
//       document.getElementById("statusMessage").className = "status";
//       document.getElementById("downloadBtn").disabled = false;

//       // Store the complete converted data for download
//       window.convertedDataForDownload = convertedData;
//     } catch (error) {
//       document.getElementById("statusMessage").textContent = "Error processing file: " + error.message;
//       document.getElementById("statusMessage").className = "status error";
//       console.error(error);
//     }
//   };
//   reader.readAsArrayBuffer(onboardingFile);
// }

// // Function to convert text to uppercase and format numbers
// function formatCellValue(value) {
//   if (value === null || value === undefined) return "";

//   // Convert numbers to fixed decimal 0 format
//   if (typeof value === "number") {
//     return value.toFixed(0);
//   }

//   // Convert strings to uppercase
//   if (typeof value === "string") {
//     // Don't convert dates to uppercase
//     if (value.match(/\d{1,2}\/\d{1,2}\/\d{4}/)) {
//       return value;
//     }
//     return value.toUpperCase().trim();
//   }

//   return value;
// }

// // Function to generate the next employee ID for a branch
// function getNextEmployeeId(branch) {
//   if (!branchEmployeeIds[branch] || branchEmployeeIds[branch].length === 0) {
//     // If no employees for this branch, return the appropriate starting ID
//     if (branch === "BIKANER") {
//       return "BFTF0001";
//     } else if (branch === "JAGATPURA") {
//       return "BFINF0001";
//     }
//     return "NOT FOUND";
//   }

//   // Filter and get only the relevant IDs for this branch
//   let relevantIds = [];
//   if (branch === "BIKANER") {
//     relevantIds = branchEmployeeIds[branch].filter(id => id.startsWith("BFTF"));
//   } else if (branch === "JAGATPURA") {
//     relevantIds = branchEmployeeIds[branch].filter(id => id.startsWith("BFINF"));
//   } else {
//     relevantIds = branchEmployeeIds[branch];
//   }

//   // If no relevant IDs found (all were wrong prefix), start from 1
//   if (relevantIds.length === 0) {
//     if (branch === "BIKANER") {
//       return "BFTF0001";
//     } else if (branch === "JAGATPURA") {
//       return "BFINF0001";
//     }
//     return "NOT FOUND";
//   }

//   // Get the last relevant ID
//   const lastId = relevantIds[relevantIds.length - 1];

//   // Extract numeric part
//   let prefix;
//   if (branch === "BIKANER") {
//     prefix = "BFTF";
//   } else if (branch === "JAGATPURA") {
//     prefix = "BFINF";
//   } else {
//     prefix = lastId.match(/^[A-Za-z]+/)[0];
//   }
  
//   const numStr = lastId.match(/\d+$/)[0];
//   const num = parseInt(numStr);

//   // Generate next ID
//   const nextNum = num + 1;
//   const nextNumStr = nextNum.toString().padStart(numStr.length, "0");

//   return prefix + nextNumStr;
// }

// function formatDateToDDMMYYYY(dateStr) {
//   if (!dateStr) return '';
  
//   // Handle cases where date might be a Date object
//   if (dateStr instanceof Date) {
//     const day = dateStr.getDate().toString().padStart(2, '0');
//     const month = (dateStr.getMonth() + 1).toString().padStart(2, '0');
//     const year = dateStr.getFullYear();
//     return `${day}/${month}/${year}`;
//   }

//   // Convert to string if it isn't already
//   dateStr = dateStr.toString().trim();

//   // Handle Excel serial date numbers (if needed)
//   if (/^\d+$/.test(dateStr)) {
//     const date = new Date((parseInt(dateStr) - (25567 + 2)) * 86400 * 1000);
//     const day = date.getDate().toString().padStart(2, '0');
//     const month = (date.getMonth() + 1).toString().padStart(2, '0');
//     const year = date.getFullYear();
//     return `${day}/${month}/${year}`;
//   }

//   // Extract numbers - handle various separators (/,-,., etc.)
//   const dateParts = dateStr.split(/[\/\-\.]/);
  
//   let day, month, year;
  
//   if (dateParts.length >= 3) {
//     // Determine format (mm/dd/yyyy vs dd/mm/yyyy)
//     if (dateParts[0].length === 4) {
//       // yyyy-mm-dd format (ISO)
//       year = dateParts[0];
//       month = dateParts[1];
//       day = dateParts[2];
//     } else if (parseInt(dateParts[0]) > 12 && parseInt(dateParts[1]) <= 12) {
//       // dd-mm-yyyy format (day first when unambiguous)
//       day = dateParts[0];
//       month = dateParts[1];
//       year = dateParts[2];
//     } else {
//       // Try to handle ambiguous cases (like 01-02-2023)
//       // Default to first part being day if it's > 31 (invalid day)
//       if (parseInt(dateParts[0]) > 31) {
//         // Probably yyyy-mm-dd
//         year = dateParts[0];
//         month = dateParts[1];
//         day = dateParts[2];
//       } else if (parseInt(dateParts[1]) > 12) {
//         // Probably mm-dd-yyyy (second part is >12 so must be day)
//         month = dateParts[0];
//         day = dateParts[1];
//         year = dateParts[2];
//       } else {
//         // Ambiguous, assume dd-mm-yyyy
//         day = dateParts[0];
//         month = dateParts[1];
//         year = dateParts[2];
//       }
//     }
    
//     // Clean up year (handle 2-digit years)
//     if (year.length === 2) {
//       year = parseInt(year) < 30 ? `20${year}` : `19${year}`;
//     }
    
//     // Pad month and day with leading zeros if needed
//     day = day.padStart(2, '0');
//     month = month.padStart(2, '0');
    
//     return `${day}/${month}/${year}`;
//   }
  
//   // Return original if we can't parse it
//   return dateStr;
// }

// // Function to convert onboarding data to employee format
// function convertOnboardingData(inputData) {
//   if (inputData.length < 2) return { employeeData: [], personalData: [] };

//   // Get headers from input sheet
//   const inputHeaders = inputData[0].map((h) =>
//     typeof h === "string" ? h.toUpperCase() : h
//   );

//   // Prepare Employee Data sheet (first sheet)
//   const employeeHeaders = [
//     "Employee Id", "Gender", "Title", "First Name", "Middile Name", "Last Name",
//     "Father Name", "Mother Name", "Spouse Name", "DOB", "DOJ", "DOC Days", "DOC",
//     "Notice Days", "DOL", "RptHead1", "RptHead2", "Location", "Designation",
//     "Department", "Branch", "Project", "Division", "Category", "CostCenter",
//     "Grade", "Attn Code", "Aadhar No", "Aadhar Name", "PF Apply", "PF UAN No",
//     "Emp Name in PF Acc", "PF Apply Date", "ESIC Apply", "ESIC No",
//     "Emp Name in ESIC Acc", "ESIC Apply Date", "PAN Name", "PAN No", "PT Apply",
//     "PT Master", "PayMode", "Bank Name", "Bank Account", "Emp Name in Bank",
//     "IFSC Code", "Temporary", "Structure Name", "Security Applicable",
//     "Security Applicable Till", "PF Number", "OT Apply", "SHL", "CO Apply",
//     "NPS Number", "CPF Number", "GPF Number", "TDS Regime", "TDS Deducted"
//   ];

//   // Prepare Personal Data sheet (second sheet)
//   const personalHeaders = [
//     "Employee Id", "Contact Num1", "Contact Num2", "Email", "Official Email",
//     "Altr. Email", "Pr_Address1", "Pr_Address2", "Pr_City", "Pr_State",
//     "Pr_Country", "Pr_PinCode", "Pm_Address1", "Pm_Address2", "Pm_City",
//     "Pm_State", "Pm_Country", "Pm_PinCode", "Emer.Contact", "Official Contact",
//     "Marriage Status", "Marriage Date", "Blood Group", "Birth Mark", "Religion",
//     "Caste Category", "DL Name", "License No", "Issue Date", "Expiry Date"
//   ];

//   const employeeData = [employeeHeaders];
//   const personalData = [personalHeaders];

//   // Process each row (skip header row)
//   for (let i = 1; i < inputData.length; i++) {
//     const row = inputData[i];
//     if (row.length === 0) continue;

//     // Get common values from input
//     const storeName = formatCellValue(getValue(row, inputHeaders, "STORE NAME"));
//     const joiningDate = formatCellValue(getValue(row, inputHeaders, "JOINING DATE"));
//     const designation = formatCellValue(getValue(row, inputHeaders, "DESIGNATION"));
//     const fullName = formatCellValue(getValue(row, inputHeaders, "FULL NAME (AS PER AADHAR CARD)"));
//     const fathersName = formatCellValue(getValue(row, inputHeaders, "FATHER'S NAME"));
//     const mothersName = formatCellValue(getValue(row, inputHeaders, "MOTHER'S NAME"));
//     const dob = formatCellValue(getValue(row, inputHeaders, "DOB(DATE OF BIRTH )"));
//     const aadharNo = formatCellValue(getValue(row, inputHeaders, "AADHAR CARD NO"));
//     const panNo = formatCellValue(getValue(row, inputHeaders, "PANCARD NO"));
//     const bankName = formatCellValue(getValue(row, inputHeaders, "BANK NAME"));
//     const bankAccount = formatCellValue(getValue(row, inputHeaders, "BANK DETAILS (ACCOUNT NO )"));
//     const ifscCode = formatCellValue(getValue(row, inputHeaders, "BANK IFSC CODE"));
//     const contactNum1 = formatCellValue(getValue(row, inputHeaders, "PERSONAL CONTACT NO"));
//     const email = formatCellValue(getValue(row, inputHeaders, "EMAIL ID"));
//     const emergencyContact = formatCellValue(getValue(row, inputHeaders, "EMERGENCY CONTACT NO FAMILY (WITH RELATION)"));
//     const bloodGroup = formatCellValue(getValue(row, inputHeaders, "BLOOD GROUP"));
//     const prAddress1 = formatCellValue(getValue(row, inputHeaders, "PRESENT ADDRESS"));
//     const pmAddress1 = formatCellValue(getValue(row, inputHeaders, "PERMANENT ADDRESS"));
//     const maritalStatus = formatCellValue(getValue(row, inputHeaders, "MARITAL STATUS"));

//     // Get branch details from the branchDetailsMap
//     const branchDetails = branchDetailsMap[storeName] || {};
//     // const location = branchDetails.location || storeName;
//     const division = branchDetails.division || "";
//     const location = storeName+"-"+division;
//     // Generate next employee ID for this branch
//     const nextEmployeeId = getNextEmployeeId(storeName);

//     // Employee Data sheet row
//     const employeeRow = [
//       nextEmployeeId, "", "", fullName || "", "", "", 
//       fathersName, mothersName, "", formatDateToDDMMYYYY(dob), formatDateToDDMMYYYY(joiningDate),
//       "180", calculateDOC(formatDateToDDMMYYYY(joiningDate)), "30", "", "SELF", "", 
//       location, designation, "", storeName, "", division, "", "", "", 
//       nextEmployeeId, aadharNo, fullName, "True", "", fullName,
//       formatDateToDDMMYYYY(joiningDate), "True", "", fullName, formatDateToDDMMYYYY(joiningDate), fullName,
//       panNo, "", "","NEFT",bankName, bankAccount, fullName, ifscCode,
//       "", "STRUCTURE 1", "", "", "", "", "", "", "", "", "", "", ""
//     ];

//     // Personal Data sheet row (only filling specified columns)
//     const personalRow = [
//       nextEmployeeId, // Employee Id
//       contactNum1, // Contact Num1 (from Google form)
//       emergencyContact, // Contact Num2 (empty)
//       email, // Email (from Google form)
//       "", // Official Email (empty)
//       "", // Altr. Email (empty)
//       prAddress1, // Pr_Address1 (from Google form)
//       "", // Pr_Address2 (empty)
//       "", // Pr_City (empty)
//       "", // Pr_State (empty)
//       "", // Pr_Country (empty)
//       "", // Pr_PinCode (empty)
//       pmAddress1, // Pm_Address1 (from Google form)
//       "", // Pm_Address2 (empty)
//       "", // Pm_City (empty)
//       "", // Pm_State (empty)
//       "", // Pm_Country (empty)
//       "", // Pm_PinCode (empty)
//       emergencyContact, // Emer.Contact (from Google form)
//       "", // Official Contact (empty)
//       maritalStatus, // Marriage Status (from Google form)
//       "", // Marriage Date (empty)
//       bloodGroup, // Blood Group (from Google form)
//       "", // Birth Mark (empty)
//       "", // Religion (empty)
//       "", // Caste Category (empty)
//       "", // DL Name (empty)
//       "", // License No (empty)
//       "", // Issue Date (empty)
//       "", // Expiry Date (empty)
//     ];

//     employeeData.push(employeeRow);
//     personalData.push(personalRow);

//     // Update the branchEmployeeIds with the newly generated ID
//     if (!branchEmployeeIds[storeName]) {
//       branchEmployeeIds[storeName] = [];
//     }
//     branchEmployeeIds[storeName].push(nextEmployeeId);
//   }

//   return {
//     employeeData: employeeData,
//     personalData: personalData
//   };
// }

// // Helper function to get value from row based on header
// function getValue(row, headers, headerName) {
//   const index = headers.indexOf(headerName.toUpperCase());
//   return index !== -1 && row[index] !== undefined ? row[index] : "";
// }

// // Helper function to format date
// function formatDate(dateStr) {
//   if (!dateStr) return "";

//   // Handle different date formats
//   if (typeof dateStr === "string" && dateStr.includes("-")) {
//     const parts = dateStr.split("-");
//     if (parts.length >= 3) {
//       return `${parts[2].substr(0, 2)}/${parts[1]}/${parts[0]}`;
//     }
//   }

//   // Handle Excel date numbers (if needed)
//   if (typeof dateStr === "number") {
//     const date = new Date((dateStr - (25567 + 2)) * 86400 * 1000);
//     return `${date.getDate()}/${date.getMonth() + 1}/${date.getFullYear()}`;
//   }

//   // Return as-is if we can't parse it
//   return dateStr;
// }

// // Calculate DOC (Date of Joining + 180 days)
// function calculateDOC(joiningDateStr) {
//   if (!joiningDateStr || typeof joiningDateStr !== "string") return "";

//   try {
//     let day, month, year;
//     let dateParts;

//     // Normalize separators to '-'
//     joiningDateStr = joiningDateStr.replace(/[^0-9]/g, "-"); // Replace all non-digits with -
//     dateParts = joiningDateStr.split("-").filter(part => part !== "");

//     if (dateParts.length !== 3) return "";

//     // Parse all parts as integers
//     const part1 = parseInt(dateParts[0], 10);
//     const part2 = parseInt(dateParts[1], 10);
//     const part3 = parseInt(dateParts[2], 10);

//     // Detect format
//     if (dateParts[0].length === 4) {
//       // YYYY-MM-DD format
//       year = part1;
//       month = part2;
//       day = part3;
//     } else if (dateParts[2].length === 4) {
//       // DD-MM-YYYY or MM-DD-YYYY format
//       year = part3;
      
//       if (part1 > 12) {
//         // Definitely DD-MM-YYYY (day > 12)
//         day = part1;
//         month = part2;
//       } else if (part2 > 12) {
//         // Definitely MM-DD-YYYY (day > 12)
//         month = part1;
//         day = part2;
//       } else if (part1 > 31) {
//         // Invalid day
//         return "";
//       } else if (part2 > 31) {
//         // Invalid day
//         return "";
//       } else {
//         // Ambiguous case (both <=12) - use DD-MM-YYYY as default
//         day = part1;
//         month = part2;
//       }
//     } else {
//       return "";
//     }

//     // Validate year range
//     if (year < 1000 || year > 9999) return "";

//     // Validate month
//     if (month < 1 || month > 12) return "";

//     // Create date object (using UTC to avoid timezone issues)
//     const joiningDate = new Date(Date.UTC(year, month - 1, day));
//     if (isNaN(joiningDate.getTime())) return "";

//     // Verify the parsed date matches the input (to catch invalid dates like Feb 30)
//     if (joiningDate.getUTCFullYear() !== year || 
//         joiningDate.getUTCMonth() + 1 !== month || 
//         joiningDate.getUTCDate() !== day) {
//       return "";
//     }

//     // Add 180 days
//     const docDate = new Date(joiningDate);
//     docDate.setUTCDate(docDate.getUTCDate() + 180);

//     // Format as DD/MM/YYYY with leading zeros
//     const format = (num) => String(num).padStart(2, "0");
//     return `${format(docDate.getUTCDate())}/${format(docDate.getUTCMonth() + 1)}/${docDate.getUTCFullYear()}`;
//   } catch (e) {
//     console.error("Error calculating DOC:", e);
//     return "";
//   }
// }

// // Display preview of converted data
// function displayPreview(data) {
//   const container = document.getElementById("tableContainer");
//   container.innerHTML = "";

//   if (data.length === 0) {
//     container.innerHTML = "<p>No data to display</p>";
//     return;
//   }

//   const table = document.createElement("table");

//   // Create header
//   const thead = document.createElement("thead");
//   const headerRow = document.createElement("tr");
//   data[0].forEach((header) => {
//     const th = document.createElement("th");
//     th.textContent = header;
//     headerRow.appendChild(th);
//   });
//   thead.appendChild(headerRow);
//   table.appendChild(thead);

//   // Create body (limit to 5 rows for preview)
//   const tbody = document.createElement("tbody");
//   const rowCount = Math.min(data.length, 6); // Show up to 5 data rows + header
//   for (let i = 1; i < rowCount; i++) {
//     const row = document.createElement("tr");
//     data[i].forEach((cell) => {
//       const td = document.createElement("td");
//       td.textContent = cell;
//       row.appendChild(td);
//     });
//     tbody.appendChild(row);
//   }
//   table.appendChild(tbody);

//   container.appendChild(table);

//   if (data.length > 6) {
//     const moreText = document.createElement("p");
//     moreText.textContent = `...and ${
//       data.length - 6
//     } more rows (not shown in preview)`;
//     container.appendChild(moreText);
//   }
// }

// // Update download button handler to include both sheets
// document.getElementById("downloadBtn").addEventListener("click", function () {
//   if (!window.convertedDataForDownload) return;

//   // Create a new workbook
//   const wb = XLSX.utils.book_new();
  
//   // Add Employee Data sheet
//   const employeeWs = XLSX.utils.aoa_to_sheet(window.convertedDataForDownload.employeeData);
//   XLSX.utils.book_append_sheet(wb, employeeWs, "Employee");
  
//   // Add Personal Data sheet
//   const personalWs = XLSX.utils.aoa_to_sheet(window.convertedDataForDownload.personalData);
//   XLSX.utils.book_append_sheet(wb, personalWs, "Personal");

//   // Generate the file and trigger download as .xlsx
//   XLSX.writeFile(wb, fileName, { bookType: "xlsx" });
// });

// // Process employee ID data and build the mapping
// function processEmployeeIdData(data) {
//   employeeIdMap = {};
//   branchEmployeeIds = {};

//   if (data.length < 2) return;

//   // Get headers
//   const headers = data[0].map((h) =>
//     typeof h === "string" ? h.toUpperCase() : h
//   );
//   const idIndex = headers.indexOf("EMPLOYEE ID");
//   const branchIndex = headers.indexOf("BRANCH");

//   // Process each row
//   for (let i = 1; i < data.length; i++) {
//     const row = data[i];
//     if (row.length <= Math.max(idIndex, branchIndex)) continue;

//     const employeeId = formatCellValue(row[idIndex]);
//     const branch = formatCellValue(row[branchIndex]);

//     // Store mapping of employee ID to branch
//     employeeIdMap[employeeId] = branch;

//     // Group employee IDs by branch
//     if (!branchEmployeeIds[branch]) {
//       branchEmployeeIds[branch] = [];
//     }
//     branchEmployeeIds[branch].push(employeeId);
//   }

//   // Sort employee IDs within each branch to find the next ID
//   for (const branch in branchEmployeeIds) {
//     branchEmployeeIds[branch].sort((a, b) => {
//       // Extract numeric parts
//       const numA = parseInt(a.replace(/^\D+/g, ""));
//       const numB = parseInt(b.replace(/^\D+/g, ""));
//       return numA - numB;
//     });
//   }
// }


// Global variables
let convertedData = [];
let fileName = "";
let employeeIdMap = {};
let branchEmployeeIds = {};
let branchDetailsMap = {};

// Handle onboarding file input change
document
  .getElementById("onboardingFileInput")
  .addEventListener("change", function (e) {
    const file = e.target.files[0];
    if (!file) return;

    document.getElementById(
      "onboardingFileName"
    ).textContent = `Selected: ${file.name}`;
    document.getElementById("statusMessage").textContent =
      "Please upload all files to proceed";
    document.getElementById("statusMessage").className = "status";
  });

// Handle employee ID file input change
document
  .getElementById("employeeIdFileInput")
  .addEventListener("change", function (e) {
    const file = e.target.files[0];
    if (!file) return;

    document.getElementById(
      "employeeIdFileName"
    ).textContent = `Selected: ${file.name}`;
    document.getElementById("statusMessage").textContent =
      "Processing employee ID file...";
    document.getElementById("statusMessage").className = "status";

    const reader = new FileReader();
    reader.onload = function (e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });

        // Assuming the first sheet is the one we want
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        // Convert to JSON
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        // Process the employee ID data
        processEmployeeIdData(jsonData);

        document.getElementById("statusMessage").textContent =
          "Employee ID file processed. Please upload other files if not done yet.";
        document.getElementById("statusMessage").className = "status";

        // Check if all files are uploaded to process
        if (areAllFilesUploaded()) {
          processAllFiles();
        }
      } catch (error) {
        document.getElementById("statusMessage").textContent =
          "Error processing employee ID file: " + error.message;
        document.getElementById("statusMessage").className = "status error";
        console.error(error);
      }
    };
    reader.readAsArrayBuffer(file);
  });

// Handle branch details file input change
document
  .getElementById("branchDetailsFileInput")
  .addEventListener("change", function (e) {
    const file = e.target.files[0];
    if (!file) return;

    document.getElementById(
      "branchDetailsFileName"
    ).textContent = `Selected: ${file.name}`;
    document.getElementById("statusMessage").textContent =
      "Processing branch details file...";
    document.getElementById("statusMessage").className = "status";

    const reader = new FileReader();
    reader.onload = function (e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });

        // Assuming the first sheet is the one we want
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        // Convert to JSON
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        // Process the branch details data
        processBranchDetailsData(jsonData);

        document.getElementById("statusMessage").textContent =
          "Branch details file processed. Please upload other files if not done yet.";
        document.getElementById("statusMessage").className = "status";

        // Check if all files are uploaded to process
        if (areAllFilesUploaded()) {
          processAllFiles();
        }
      } catch (error) {
        document.getElementById("statusMessage").textContent =
          "Error processing branch details file: " + error.message;
        document.getElementById("statusMessage").className = "status error";
        console.error(error);
      }
    };
    reader.readAsArrayBuffer(file);
  });

// Process branch details data and build the mapping
function processBranchDetailsData(data) {
  branchDetailsMap = {};

  if (data.length < 2) return;

  // Get headers
  const headers = data[0].map((h) =>
    typeof h === "string" ? h.toUpperCase() : h
  );
  const branchIndex = headers.indexOf("BRANCH");
  const locationIndex = headers.indexOf("LOCATION");
  const divisionIndex = headers.indexOf("DIVISION");
  const RptHead1Index = headers.indexOf("RPTHEAD1");

  // Process each row
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row.length <= Math.max(branchIndex, locationIndex, divisionIndex, RptHead1Index)) continue;

    const branch = formatCellValue(row[branchIndex]);
    const location = formatCellValue(row[locationIndex]);
    const division = formatCellValue(row[divisionIndex]);
    const RptHead1 = formatCellValue(row[RptHead1Index]);

    // Store mapping of branch to location, division and RptHead1
    if (branch) {
      branchDetailsMap[branch] = {
        location: location || "",
        division: division || "",
        RptHead1: RptHead1 || "SELF" // Default to "SELF" if not provided
      };
    }
  }
}

// Check if all required files are uploaded
function areAllFilesUploaded() {
  return (
    document.getElementById("onboardingFileInput").files.length > 0 &&
    document.getElementById("employeeIdFileInput").files.length > 0 &&
    document.getElementById("branchDetailsFileInput").files.length > 0
  );
}

// Process all files when all are uploaded
function processAllFiles() {
  const onboardingFile = document.getElementById("onboardingFileInput").files[0];
  if (!onboardingFile) return;

  fileName = onboardingFile.name.replace(/\.[^/.]+$/, "") + "_converted.xlsx";
  document.getElementById("statusMessage").textContent = "Processing all files...";
  document.getElementById("statusMessage").className = "status";

  const reader = new FileReader();
  reader.onload = function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });

      // Assuming the first sheet is the one we want
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];

      // Convert to JSON
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      // Process the data (now returns an object with both sheets)
      const convertedData = convertOnboardingData(jsonData);

      // Display preview (showing just the first sheet)
      displayPreview(convertedData.employeeData);

      document.getElementById("statusMessage").textContent = "Files processed successfully!";
      document.getElementById("statusMessage").className = "status";
      document.getElementById("downloadBtn").disabled = false;

      // Store the complete converted data for download
      window.convertedDataForDownload = convertedData;
    } catch (error) {
      document.getElementById("statusMessage").textContent = "Error processing file: " + error.message;
      document.getElementById("statusMessage").className = "status error";
      console.error(error);
    }
  };
  reader.readAsArrayBuffer(onboardingFile);
}

// Function to convert text to uppercase and format numbers
function formatCellValue(value) {
  if (value === null || value === undefined) return "";

  // Convert numbers to fixed decimal 0 format
  if (typeof value === "number") {
    return value.toFixed(0);
  }

  // Convert strings to uppercase
  if (typeof value === "string") {
    // Don't convert dates to uppercase
    if (value.match(/\d{1,2}\/\d{1,2}\/\d{4}/)) {
      return value;
    }
    return value.toUpperCase().trim();
  }

  return value;
}

// Function to generate the next employee ID for a branch
function getNextEmployeeId(branch) {
  if (!branchEmployeeIds[branch] || branchEmployeeIds[branch].length === 0) {
    // If no employees for this branch, return the appropriate starting ID
    if (branch === "BIKANER") {
      return "BFTF0001";
    } else if (branch === "JAGATPURA") {
      return "BFINF0001";
    }
    return "NOT FOUND";
  }

  // Filter and get only the relevant IDs for this branch
  let relevantIds = [];
  if (branch === "BIKANER") {
    relevantIds = branchEmployeeIds[branch].filter(id => id.startsWith("BFTF"));
  } else if (branch === "JAGATPURA") {
    relevantIds = branchEmployeeIds[branch].filter(id => id.startsWith("BFINF"));
  } else {
    relevantIds = branchEmployeeIds[branch];
  }

  // If no relevant IDs found (all were wrong prefix), start from 1
  if (relevantIds.length === 0) {
    if (branch === "BIKANER") {
      return "BFTF0001";
    } else if (branch === "JAGATPURA") {
      return "BFINF0001";
    }
    return "NOT FOUND";
  }

  // Get the last relevant ID
  const lastId = relevantIds[relevantIds.length - 1];

  // Extract numeric part
  let prefix;
  if (branch === "BIKANER") {
    prefix = "BFTF";
  } else if (branch === "JAGATPURA") {
    prefix = "BFINF";
  } else {
    prefix = lastId.match(/^[A-Za-z]+/)[0];
  }
  
  const numStr = lastId.match(/\d+$/)[0];
  const num = parseInt(numStr);

  // Generate next ID
  const nextNum = num + 1;
  const nextNumStr = nextNum.toString().padStart(numStr.length, "0");

  return prefix + nextNumStr;
}

function formatDateToDDMMYYYY(dateStr) {
  if (!dateStr) return '';
  
  // Handle cases where date might be a Date object
  if (dateStr instanceof Date) {
    const day = dateStr.getDate().toString().padStart(2, '0');
    const month = (dateStr.getMonth() + 1).toString().padStart(2, '0');
    const year = dateStr.getFullYear();
    return `${day}/${month}/${year}`;
  }

  // Convert to string if it isn't already
  dateStr = dateStr.toString().trim();

  // Handle Excel serial date numbers (if needed)
  if (/^\d+$/.test(dateStr)) {
    const date = new Date((parseInt(dateStr) - (25567 + 2)) * 86400 * 1000);
    const day = date.getDate().toString().padStart(2, '0');
    const month = (date.getMonth() + 1).toString().padStart(2, '0');
    const year = date.getFullYear();
    return `${day}/${month}/${year}`;
  }

  // Extract numbers - handle various separators (/,-,., etc.)
  const dateParts = dateStr.split(/[\/\-\.]/);
  
  let day, month, year;
  
  if (dateParts.length >= 3) {
    // Determine format (mm/dd/yyyy vs dd/mm/yyyy)
    if (dateParts[0].length === 4) {
      // yyyy-mm-dd format (ISO)
      year = dateParts[0];
      month = dateParts[1];
      day = dateParts[2];
    } else if (parseInt(dateParts[0]) > 12 && parseInt(dateParts[1]) <= 12) {
      // dd-mm-yyyy format (day first when unambiguous)
      day = dateParts[0];
      month = dateParts[1];
      year = dateParts[2];
    } else {
      // Try to handle ambiguous cases (like 01-02-2023)
      // Default to first part being day if it's > 31 (invalid day)
      if (parseInt(dateParts[0]) > 31) {
        // Probably yyyy-mm-dd
        year = dateParts[0];
        month = dateParts[1];
        day = dateParts[2];
      } else if (parseInt(dateParts[1]) > 12) {
        // Probably mm-dd-yyyy (second part is >12 so must be day)
        month = dateParts[0];
        day = dateParts[1];
        year = dateParts[2];
      } else {
        // Ambiguous, assume dd-mm-yyyy
        day = dateParts[0];
        month = dateParts[1];
        year = dateParts[2];
      }
    }
    
    // Clean up year (handle 2-digit years)
    if (year.length === 2) {
      year = parseInt(year) < 30 ? `20${year}` : `19${year}`;
    }
    
    // Pad month and day with leading zeros if needed
    day = day.padStart(2, '0');
    month = month.padStart(2, '0');
    
    return `${day}/${month}/${year}`;
  }
  
  // Return original if we can't parse it
  return dateStr;
}

// Function to convert onboarding data to employee format
function convertOnboardingData(inputData) {
  if (inputData.length < 2) return { employeeData: [], personalData: [] };

  // Get headers from input sheet
  const inputHeaders = inputData[0].map((h) =>
    typeof h === "string" ? h.toUpperCase() : h
  );

  // Prepare Employee Data sheet (first sheet)
  const employeeHeaders = [
    "Employee Id", "Gender", "Title", "First Name", "Middile Name", "Last Name",
    "Father Name", "Mother Name", "Spouse Name", "DOB", "DOJ", "DOC Days", "DOC",
    "Notice Days", "DOL", "RptHead1", "RptHead2", "Location", "Designation",
    "Department", "Branch", "Project", "Division", "Category", "CostCenter",
    "Grade", "Attn Code", "Aadhar No", "Aadhar Name", "PF Apply", "PF UAN No",
    "Emp Name in PF Acc", "PF Apply Date", "ESIC Apply", "ESIC No",
    "Emp Name in ESIC Acc", "ESIC Apply Date", "PAN Name", "PAN No", "PT Apply",
    "PT Master", "PayMode", "Bank Name", "Bank Account", "Emp Name in Bank",
    "IFSC Code", "Temporary", "Structure Name", "Security Applicable",
    "Security Applicable Till", "PF Number", "OT Apply", "SHL", "CO Apply",
    "NPS Number", "CPF Number", "GPF Number", "TDS Regime", "TDS Deducted"
  ];

  // Prepare Personal Data sheet (second sheet)
  const personalHeaders = [
    "Employee Id", "Contact Num1", "Contact Num2", "Email", "Official Email",
    "Altr. Email", "Pr_Address1", "Pr_Address2", "Pr_City", "Pr_State",
    "Pr_Country", "Pr_PinCode", "Pm_Address1", "Pm_Address2", "Pm_City",
    "Pm_State", "Pm_Country", "Pm_PinCode", "Emer.Contact", "Official Contact",
    "Marriage Status", "Marriage Date", "Blood Group", "Birth Mark", "Religion",
    "Caste Category", "DL Name", "License No", "Issue Date", "Expiry Date"
  ];

  const employeeData = [employeeHeaders];
  const personalData = [personalHeaders];

  // Process each row (skip header row)
  for (let i = 1; i < inputData.length; i++) {
    const row = inputData[i];
    if (row.length === 0) continue;

    // Get common values from input
    const storeName = formatCellValue(getValue(row, inputHeaders, "STORE NAME"));
    const joiningDate = formatCellValue(getValue(row, inputHeaders, "JOINING DATE"));
    const designation = formatCellValue(getValue(row, inputHeaders, "DESIGNATION"));
    const fullName = formatCellValue(getValue(row, inputHeaders, "FULL NAME (AS PER AADHAR CARD)"));
    const fathersName = formatCellValue(getValue(row, inputHeaders, "FATHER'S NAME"));
    const mothersName = formatCellValue(getValue(row, inputHeaders, "MOTHER'S NAME"));
    const dob = formatCellValue(getValue(row, inputHeaders, "DOB(DATE OF BIRTH )"));
    const aadharNo = formatCellValue(getValue(row, inputHeaders, "AADHAR CARD NO"));
    const panNo = formatCellValue(getValue(row, inputHeaders, "PANCARD NO"));
    const bankName = formatCellValue(getValue(row, inputHeaders, "BANK NAME"));
    const bankAccount = formatCellValue(getValue(row, inputHeaders, "BANK DETAILS (ACCOUNT NO )"));
    const ifscCode = formatCellValue(getValue(row, inputHeaders, "BANK IFSC CODE"));
    const contactNum1 = formatCellValue(getValue(row, inputHeaders, "PERSONAL CONTACT NO"));
    const email = formatCellValue(getValue(row, inputHeaders, "EMAIL ID"));
    const emergencyContact = formatCellValue(getValue(row, inputHeaders, "EMERGENCY CONTACT NO FAMILY (WITH RELATION)"));
    const bloodGroup = formatCellValue(getValue(row, inputHeaders, "BLOOD GROUP"));
    const prAddress1 = formatCellValue(getValue(row, inputHeaders, "PRESENT ADDRESS"));
    const pmAddress1 = formatCellValue(getValue(row, inputHeaders, "PERMANENT ADDRESS"));
    const maritalStatus = formatCellValue(getValue(row, inputHeaders, "MARITAL STATUS"));

    // Get branch details from the branchDetailsMap
    const branchDetails = branchDetailsMap[storeName] || {};
    const division = branchDetails.division || "";
    const location = storeName+"-"+division;
    const RptHead1 = branchDetails.RptHead1 || "SELF"; // Use RptHead1 from branch details or default to "SELF"

    // Generate next employee ID for this branch
    const nextEmployeeId = getNextEmployeeId(storeName);

    // Employee Data sheet row
    const employeeRow = [
      nextEmployeeId, "", "", fullName || "", "", "", 
      fathersName, mothersName, "", formatDateToDDMMYYYY(dob), formatDateToDDMMYYYY(joiningDate),
      "180", calculateDOC(formatDateToDDMMYYYY(joiningDate)), "30", "", RptHead1, "", 
      location, designation, "", storeName, "", division, "", "", "", 
      nextEmployeeId, aadharNo, fullName, "True", "", fullName,
      formatDateToDDMMYYYY(joiningDate), "True", "", fullName, formatDateToDDMMYYYY(joiningDate), fullName,
      panNo, "", "","NEFT",bankName, bankAccount, fullName, ifscCode,
      "", "STRUCTURE 1", "", "", "", "", "", "", "", "", "", "", ""
    ];

    // Personal Data sheet row (only filling specified columns)
    const personalRow = [
      nextEmployeeId, // Employee Id
      contactNum1, // Contact Num1 (from Google form)
      emergencyContact, // Contact Num2 (empty)
      email, // Email (from Google form)
      "", // Official Email (empty)
      "", // Altr. Email (empty)
      prAddress1, // Pr_Address1 (from Google form)
      "", // Pr_Address2 (empty)
      "", // Pr_City (empty)
      "", // Pr_State (empty)
      "", // Pr_Country (empty)
      "", // Pr_PinCode (empty)
      pmAddress1, // Pm_Address1 (from Google form)
      "", // Pm_Address2 (empty)
      "", // Pm_City (empty)
      "", // Pm_State (empty)
      "", // Pm_Country (empty)
      "", // Pm_PinCode (empty)
      emergencyContact, // Emer.Contact (from Google form)
      "", // Official Contact (empty)
      maritalStatus, // Marriage Status (from Google form)
      "", // Marriage Date (empty)
      bloodGroup, // Blood Group (from Google form)
      "", // Birth Mark (empty)
      "", // Religion (empty)
      "", // Caste Category (empty)
      "", // DL Name (empty)
      "", // License No (empty)
      "", // Issue Date (empty)
      "", // Expiry Date (empty)
    ];

    employeeData.push(employeeRow);
    personalData.push(personalRow);

    // Update the branchEmployeeIds with the newly generated ID
    if (!branchEmployeeIds[storeName]) {
      branchEmployeeIds[storeName] = [];
    }
    branchEmployeeIds[storeName].push(nextEmployeeId);
  }

  return {
    employeeData: employeeData,
    personalData: personalData
  };
}

// Helper function to get value from row based on header
function getValue(row, headers, headerName) {
  const index = headers.indexOf(headerName.toUpperCase());
  return index !== -1 && row[index] !== undefined ? row[index] : "";
}

// Helper function to format date
function formatDate(dateStr) {
  if (!dateStr) return "";

  // Handle different date formats
  if (typeof dateStr === "string" && dateStr.includes("-")) {
    const parts = dateStr.split("-");
    if (parts.length >= 3) {
      return `${parts[2].substr(0, 2)}/${parts[1]}/${parts[0]}`;
    }
  }

  // Handle Excel date numbers (if needed)
  if (typeof dateStr === "number") {
    const date = new Date((dateStr - (25567 + 2)) * 86400 * 1000);
    return `${date.getDate()}/${date.getMonth() + 1}/${date.getFullYear()}`;
  }

  // Return as-is if we can't parse it
  return dateStr;
}

// Calculate DOC (Date of Joining + 180 days)
function calculateDOC(joiningDateStr) {
  if (!joiningDateStr || typeof joiningDateStr !== "string") return "";

  try {
    let day, month, year;
    let dateParts;

    // Normalize separators to '-'
    joiningDateStr = joiningDateStr.replace(/[^0-9]/g, "-"); // Replace all non-digits with -
    dateParts = joiningDateStr.split("-").filter(part => part !== "");

    if (dateParts.length !== 3) return "";

    // Parse all parts as integers
    const part1 = parseInt(dateParts[0], 10);
    const part2 = parseInt(dateParts[1], 10);
    const part3 = parseInt(dateParts[2], 10);

    // Detect format
    if (dateParts[0].length === 4) {
      // YYYY-MM-DD format
      year = part1;
      month = part2;
      day = part3;
    } else if (dateParts[2].length === 4) {
      // DD-MM-YYYY or MM-DD-YYYY format
      year = part3;
      
      if (part1 > 12) {
        // Definitely DD-MM-YYYY (day > 12)
        day = part1;
        month = part2;
      } else if (part2 > 12) {
        // Definitely MM-DD-YYYY (day > 12)
        month = part1;
        day = part2;
      } else if (part1 > 31) {
        // Invalid day
        return "";
      } else if (part2 > 31) {
        // Invalid day
        return "";
      } else {
        // Ambiguous case (both <=12) - use DD-MM-YYYY as default
        day = part1;
        month = part2;
      }
    } else {
      return "";
    }

    // Validate year range
    if (year < 1000 || year > 9999) return "";

    // Validate month
    if (month < 1 || month > 12) return "";

    // Create date object (using UTC to avoid timezone issues)
    const joiningDate = new Date(Date.UTC(year, month - 1, day));
    if (isNaN(joiningDate.getTime())) return "";

    // Verify the parsed date matches the input (to catch invalid dates like Feb 30)
    if (joiningDate.getUTCFullYear() !== year || 
        joiningDate.getUTCMonth() + 1 !== month || 
        joiningDate.getUTCDate() !== day) {
      return "";
    }

    // Add 180 days
    const docDate = new Date(joiningDate);
    docDate.setUTCDate(docDate.getUTCDate() + 180);

    // Format as DD/MM/YYYY with leading zeros
    const format = (num) => String(num).padStart(2, "0");
    return `${format(docDate.getUTCDate())}/${format(docDate.getUTCMonth() + 1)}/${docDate.getUTCFullYear()}`;
  } catch (e) {
    console.error("Error calculating DOC:", e);
    return "";
  }
}

// Display preview of converted data
function displayPreview(data) {
  const container = document.getElementById("tableContainer");
  container.innerHTML = "";

  if (data.length === 0) {
    container.innerHTML = "<p>No data to display</p>";
    return;
  }

  const table = document.createElement("table");

  // Create header
  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");
  data[0].forEach((header) => {
    const th = document.createElement("th");
    th.textContent = header;
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);
  table.appendChild(thead);

  // Create body (limit to 5 rows for preview)
  const tbody = document.createElement("tbody");
  const rowCount = Math.min(data.length, 6); // Show up to 5 data rows + header
  for (let i = 1; i < rowCount; i++) {
    const row = document.createElement("tr");
    data[i].forEach((cell) => {
      const td = document.createElement("td");
      td.textContent = cell;
      row.appendChild(td);
    });
    tbody.appendChild(row);
  }
  table.appendChild(tbody);

  container.appendChild(table);

  if (data.length > 6) {
    const moreText = document.createElement("p");
    moreText.textContent = `...and ${
      data.length - 6
    } more rows (not shown in preview)`;
    container.appendChild(moreText);
  }
}

// Update download button handler to include both sheets
document.getElementById("downloadBtn").addEventListener("click", function () {
  if (!window.convertedDataForDownload) return;

  // Create a new workbook
  const wb = XLSX.utils.book_new();
  
  // Add Employee Data sheet
  const employeeWs = XLSX.utils.aoa_to_sheet(window.convertedDataForDownload.employeeData);
  XLSX.utils.book_append_sheet(wb, employeeWs, "Employee");
  
  // Add Personal Data sheet
  const personalWs = XLSX.utils.aoa_to_sheet(window.convertedDataForDownload.personalData);
  XLSX.utils.book_append_sheet(wb, personalWs, "Personal");

  // Generate the file and trigger download as .xlsx
  XLSX.writeFile(wb, fileName, { bookType: "xlsx" });
});

// Process employee ID data and build the mapping
function processEmployeeIdData(data) {
  employeeIdMap = {};
  branchEmployeeIds = {};

  if (data.length < 2) return;

  // Get headers
  const headers = data[0].map((h) =>
    typeof h === "string" ? h.toUpperCase() : h
  );
  const idIndex = headers.indexOf("EMPLOYEE ID");
  const branchIndex = headers.indexOf("BRANCH");

  // Process each row
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row.length <= Math.max(idIndex, branchIndex)) continue;

    const employeeId = formatCellValue(row[idIndex]);
    const branch = formatCellValue(row[branchIndex]);

    // Store mapping of employee ID to branch
    employeeIdMap[employeeId] = branch;

    // Group employee IDs by branch
    if (!branchEmployeeIds[branch]) {
      branchEmployeeIds[branch] = [];
    }
    branchEmployeeIds[branch].push(employeeId);
  }

  // Sort employee IDs within each branch to find the next ID
  for (const branch in branchEmployeeIds) {
    branchEmployeeIds[branch].sort((a, b) => {
      // Extract numeric parts
      const numA = parseInt(a.replace(/^\D+/g, ""));
      const numB = parseInt(b.replace(/^\D+/g, ""));
      return numA - numB;
    });
  }
}