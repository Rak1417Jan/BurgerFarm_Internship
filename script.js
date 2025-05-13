// Global variables
let convertedData = [];
let fileName = "";
let employeeIdMap = {};
let branchEmployeeIds = {};

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
      "Please upload both files to proceed";
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
          "Employee ID file processed. Please upload onboarding file if not done yet.";
        document.getElementById("statusMessage").className = "status";

        // If onboarding file is already uploaded, process both
        if (document.getElementById("onboardingFileInput").files.length > 0) {
          processBothFiles();
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

// Process both files when both are uploaded
function processBothFiles() {
  const onboardingFile = document.getElementById("onboardingFileInput")
    .files[0];
  if (!onboardingFile) return;

  fileName = onboardingFile.name.replace(/\.[^/.]+$/, "") + "_converted.xlsx";
  document.getElementById("statusMessage").textContent = "Processing files...";
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

      // Process the data
      convertedData = convertOnboardingData(jsonData);

      // Display preview
      displayPreview(convertedData);

      document.getElementById("statusMessage").textContent =
        "Files processed successfully!";
      document.getElementById("statusMessage").className = "status";
      document.getElementById("downloadBtn").disabled = false;
    } catch (error) {
      document.getElementById("statusMessage").textContent =
        "Error processing file: " + error.message;
      document.getElementById("statusMessage").className = "status error";
      console.error(error);
    }
  };
  reader.readAsArrayBuffer(onboardingFile);
}

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
    return "BFS0001";
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
    return "BFS0001";
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

// Function to convert onboarding data to employee format
// function convertOnboardingData(inputData) {
//   if (inputData.length < 2) return [];

//   // Get headers from both sheets
//   const inputHeaders = inputData[0].map((h) =>
//     typeof h === "string" ? h.toUpperCase() : h
//   );
//   const outputHeaders = [
//     "Employee Id",
//     "Gender",
//     "Title",
//     "First Name",
//     "Middile Name",
//     "Last Name",
//     "Father Name",
//     "Mother Name",
//     "Spouse Name",
//     "DOB",
//     "DOJ",
//     "DOC Days",
//     "DOC",
//     "Notice Days",
//     "DOL",
//     "RptHead1",
//     "RptHead2",
//     "Location",
//     "Designation",
//     "Department",
//     "Branch",
//     "Project",
//     "Division",
//     "Category",
//     "CostCenter",
//     "Grade",
//     "Attn Code",
//     "Aadhar No",
//     "Aadhar Name",
//     "PF Apply",
//     "PF UAN No",
//     "Emp Name in PF Acc",
//     "PF Apply Date",
//     "ESIC Apply",
//     "ESIC No",
//     "Emp Name in ESIC Acc",
//     "ESIC Apply Date",
//     "PAN Name",
//     "PAN No",
//     "PT Apply",
//     "PT Master",
//     "PayMode",
//     "Bank Name",
//     "Bank Account",
//     "Emp Name in Bank",
//     "IFSC Code",
//     "Temporary",
//     "Structure Name",
//     "Security Applicable",
//     "Security Applicable Till",
//     "PF Number",
//     "OT Apply",
//     "SHL",
//     "CO Apply",
//     "NPS Number",
//     "CPF Number",
//     "GPF Number",
//     "TDS Regime",
//     "TDS Deducted",
//   ];

//   // Prepare output data
//   const outputData = [outputHeaders];

//   // Process each row (skip header row)
//   for (let i = 1; i < inputData.length; i++) {
//     const row = inputData[i];
//     if (row.length === 0) continue;

//     // Get values from input and format them
//     const storeName = formatCellValue(
//       getValue(row, inputHeaders, "STORE NAME")
//     );
//     const joiningDate = formatCellValue(
//       getValue(row, inputHeaders, "JOINING DATE")
//     );
//     const designation = formatCellValue(
//       getValue(row, inputHeaders, "DESIGNATION")
//     );
//     const fullName = formatCellValue(
//       getValue(row, inputHeaders, "FULL NAME (AS PER AADHAR CARD)")
//     );
//     const fathersName = formatCellValue(
//       getValue(row, inputHeaders, "FATHER'S NAME")
//     );
//     const mothersName = formatCellValue(
//       getValue(row, inputHeaders, "MOTHER'S NAME")
//     );
//     const dob = formatDate(
//       formatCellValue(getValue(row, inputHeaders, "DOB(DATE OF BIRTH )"))
//     );
//     const aadharNo = formatCellValue(
//       getValue(row, inputHeaders, "AADHAR CARD NO")
//     );
//     const panNo = formatCellValue(getValue(row, inputHeaders, "PANCARD NO"));
//     const bankName = formatCellValue(getValue(row, inputHeaders, "BANK NAME"));
//     const bankAccount = formatCellValue(
//       getValue(row, inputHeaders, "BANK DETAILS (ACCOUNT NO )")
//     );
//     const ifscCode = formatCellValue(
//       getValue(row, inputHeaders, "BANK IFSC CODE")
//     );

//     // Generate next employee ID for this branch
//     const nextEmployeeId = getNextEmployeeId(storeName);

//     // Create output row with formatted values
//     const outputRow = [
//       nextEmployeeId, // Employee Id (generated)
//       "", // Gender (Not fill)
//       "", // Title (M -> Mr / F -> Ms.) (Not fill)
//       fullName || "", // First Name (full name)
//       "", // Middile Name (empty)
//       "", // Last Name (empty)
//       fathersName, // Father Name
//       mothersName, // Mother Name
//       "", // Spouse Name (Not fill)
//       dob, // DOB
//       formatDate(joiningDate), // DOJ
//       "180", // DOC Days (always 180)
//       calculateDOC(joiningDate), // DOC (Date of Joining + 180)
//       "30", // Notice Days (always 30)
//       "", // DOL (Not fill)
//       "SELF", // RptHead1 (Self)
//       "", // RptHead2 (Not fill)
//       storeName, // Location
//       designation, // Designation
//       "", // Department (Self)
//       storeName, // Branch (Same as store name)
//       "", // Project (Not fill)
//       "", // Division (refer from sheet) (Not fill)
//       "", // Category (Not fill)
//       "", // CostCenter (Not fill)
//       "", // Grade (Not fill)
//       "", // Attn Code (Not fill)
//       aadharNo, // Aadhar No
//       fullName, // Aadhar Name (Same as emp code)
//       "", // PF Apply (Yes always) - This seems contradictory to the note
//       "", // PF UAN No (Not fill)
//       fullName, // Emp Name in PF Acc (Same as full name)
//       formatDate(joiningDate), // PF Apply Date (Same as joining date)
//       "SELF", // ESIC Apply (Self)
//       "", // ESIC No (Not fill)
//       fullName, // Emp Name in ESIC Acc (Same as full name)
//       formatDate(joiningDate), // ESIC Apply Date (Same as date of joining)
//       fullName, // PAN Name (Same as full name)
//       panNo, // PAN No
//       "SELF", // PT Apply (Self)
//       "", // PT Master (Not fill)
//       "NEFT", // PayMode (Always NEFT)
//       bankName, // Bank Name
//       bankAccount, // Bank Account
//       fullName, // Emp Name in Bank (Same as full name)
//       ifscCode, // IFSC Code
//       "", // Temporary (Not fill)
//       "STRUCTURE 1", // Structure Name (Always structure 1)
//       "", // Security Applicable (Not fill)
//       "", // Security Applicable Till (Not fill)
//       "", // PF Number (Not fill)
//       "", // OT Apply (Not fill)
//       "", // SHL (Not fill)
//       "", // CO Apply (Not fill)
//       "", // NPS Number (Not fill)
//       "", // CPF Number (Not fill)
//       "", // GPF Number (Not fill)
//       "", // TDS Regime (Not fill)
//       "", // TDS Deducted (Not fill)
//     ];

//     outputData.push(outputRow);

//     // Update the branchEmployeeIds with the newly generated ID
//     if (!branchEmployeeIds[storeName]) {
//       branchEmployeeIds[storeName] = [];
//     }
//     branchEmployeeIds[storeName].push(nextEmployeeId);
//   }

//   return outputData;
// }

// Function to convert onboarding data to employee format (updated for two sheets)
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
    const dob = formatDate(formatCellValue(getValue(row, inputHeaders, "DOB(DATE OF BIRTH )")));
    const aadharNo = formatCellValue(getValue(row, inputHeaders, "AADHAR CARD NO"));
    const panNo = formatCellValue(getValue(row, inputHeaders, "PANCARD NO"));
    const bankName = formatCellValue(getValue(row, inputHeaders, "BANK NAME"));
    const bankAccount = formatCellValue(getValue(row, inputHeaders, "BANK DETAILS (ACCOUNT NO )"));
    const ifscCode = formatCellValue(getValue(row, inputHeaders, "BANK IFSC CODE"));
    const contactNum1 = formatCellValue(getValue(row, inputHeaders, "PERSONAL CONTACT NO"));
    const email = formatCellValue(getValue(row, inputHeaders, "EMAIL ID"));
    const emergencyContact = formatCellValue(getValue(row, inputHeaders, "EMERGENCY CONTACT NO FAMILY (WITH RELATION)"));
    const bloodGroup = formatCellValue(getValue(row, inputHeaders, "BLOOD GROUP"));
    const prAddress1 = formatCellValue(getValue(row, inputHeaders, "RESIDENTIAL ADDRESS"));
    const pmAddress1 = formatCellValue(getValue(row, inputHeaders, "PERMANENT ADDRESS"));
    const maritalStatus = formatCellValue(getValue(row, inputHeaders, "MARITAL STATUS"));
    // Generate next employee ID for this branch
    const nextEmployeeId = getNextEmployeeId(storeName);

    // Employee Data sheet row
    const employeeRow = [
      nextEmployeeId, "", "", fullName || "", "", "", 
      fathersName, mothersName, "", dob, formatDate(joiningDate),
      "180", calculateDOC(joiningDate), "30", "", "SELF", "", 
      storeName, designation, "", storeName, "", "", "", "", "", 
      aadharNo, fullName, "", "", fullName, formatDate(joiningDate),
      "SELF", "", fullName, formatDate(joiningDate), fullName, panNo,
      "SELF", "", "NEFT", bankName, bankAccount, fullName, ifscCode,
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

// Update processBothFiles to handle two sheets
function processBothFiles() {
  const onboardingFile = document.getElementById("onboardingFileInput").files[0];
  if (!onboardingFile) return;

  fileName = onboardingFile.name.replace(/\.[^/.]+$/, "") + "_converted.xlsx";
  document.getElementById("statusMessage").textContent = "Processing files...";
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
  if (!joiningDateStr) return "";

  try {
    // Parse the joining date
    let dateParts;
    if (joiningDateStr.includes("/")) {
      dateParts = joiningDateStr.split("/");
    } else if (joiningDateStr.includes("-")) {
      dateParts = joiningDateStr.split("-");
    } else {
      return "";
    }

    // Create date object (handling different formats)
    let day, month, year;
    if (dateParts[0].length === 4) {
      // YYYY-MM-DD format
      year = parseInt(dateParts[0]);
      month = parseInt(dateParts[1]) - 1;
      day = parseInt(dateParts[2]);
    } else {
      // DD/MM/YYYY format
      day = parseInt(dateParts[0]);
      month = parseInt(dateParts[1]) - 1;
      year = parseInt(dateParts[2]);
    }

    const joiningDate = new Date(year, month, day);
    if (isNaN(joiningDate.getTime())) return "";

    // Add 180 days
    const docDate = new Date(joiningDate);
    docDate.setDate(joiningDate.getDate() + 180);

    // Format as DD/MM/YYYY
    return `${docDate.getDate()}/${
      docDate.getMonth() + 1
    }/${docDate.getFullYear()}`;
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

// Handle download button click
document.getElementById("downloadBtn").addEventListener("click", function () {
  if (convertedData.length === 0) return;

  // Create a new workbook
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(convertedData);

  // Add the worksheet to the workbook
  XLSX.utils.book_append_sheet(wb, ws, "Employee Data");

  // Generate the file and trigger download as .xlsx only
  XLSX.writeFile(wb, fileName, { bookType: "xlsx" });
});
