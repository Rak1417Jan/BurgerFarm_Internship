// Global variables
let employeeData = [];
let branchManagers = {}; // To store branch manager email IDs
let branchMailIDs = {}; // To store branch mail IDs
let branchInfoFile = null;
let employeeDataFile = null;

// Constants for email configuration
const CC_LIST = [
  "hr.burgerfarm@gmail.com",
  "hrburgerfarm@gmail.com",
  "hr@burgerfarm.in",
  "mahimahr.burgerfarm@gmail.com",
  "hrassociate.burgerfarm@gmail.com",
];

// DOM elements
const branchInfoInput = document.getElementById("branchInfoFile");
const branchFileName = document.getElementById("branchFileName");
const employeeDataInput = document.getElementById("employeeDataFile");
const employeeFileName = document.getElementById("employeeFileName");
const uploadBtn = document.getElementById("uploadBtn");
const loadingIndicator = document.getElementById("loadingIndicator");
const dataContainer = document.getElementById("dataContainer");
const branchDataContainer = document.getElementById("branchData");
const emailPreview = document.getElementById("emailPreview");
const emailToElement = document.getElementById("emailTo");
const emailCCElement = document.getElementById("emailCC");
const emailSubjectElement = document.getElementById("emailSubject");
const emailBodyElement = document.getElementById("emailBody");
const sendEmailBtn = document.getElementById("sendEmailBtn");
const backToListBtn = document.getElementById("backToListBtn");
const notification = document.getElementById("notification");
const notificationMessage = document.getElementById("notificationMessage");
const closeNotification = document.getElementById("closeNotification");

// Event listeners
// Event listeners
document.addEventListener("DOMContentLoaded", function () {
  branchInfoInput.addEventListener("change", handleBranchFileSelect);
  employeeDataInput.addEventListener("change", handleEmployeeFileSelect);
  uploadBtn.addEventListener("click", processFiles);
  backToListBtn.addEventListener("click", hideEmailPreview);
  closeNotification.addEventListener("click", hideNotification);

  // Add copy button event listeners
  document.addEventListener("click", function (e) {
    if (e.target.classList.contains("copy-btn")) {
      const targetId = e.target.getAttribute("data-target");
      if (targetId) {
        // Copy specific field
        const element = document.getElementById(targetId);
        copyToClipboard(element.textContent);
        showNotification(
          `${targetId.replace("email", "")} copied to clipboard!`,
          "success"
        );
      } else if (e.target.id === "copyBodyBtn") {
        // Copy email body - use the plain text version
        const plainTextBody = emailBodyElement.dataset.plainText;
        copyToClipboard(plainTextBody);
        showNotification(
          "Email body copied to clipboard as plain text!",
          "success"
        );
      }
    }
  });
});

/**
 * Copies text to clipboard
 */
function copyToClipboard(text) {
    const textarea = document.createElement('textarea');
    textarea.value = text;
    document.body.appendChild(textarea);
    textarea.select();
    
    try {
        document.execCommand('copy');
    } catch (err) {
        console.error('Failed to copy text: ', err);
    }
    
    document.body.removeChild(textarea);
}

/**
 * Handles the branch file selection event
 */
function handleBranchFileSelect(e) {
  const file = e.target.files[0];
  if (file) {
    branchInfoFile = file;
    branchFileName.textContent = file.name;
    checkFilesReady();
  } else {
    branchInfoFile = null;
    branchFileName.textContent = "No file selected";
    checkFilesReady();
  }
}

/**
 * Handles the employee file selection event
 */
function handleEmployeeFileSelect(e) {
  const file = e.target.files[0];
  if (file) {
    employeeDataFile = file;
    employeeFileName.textContent = file.name;
    checkFilesReady();
  } else {
    employeeDataFile = null;
    employeeFileName.textContent = "No file selected";
    checkFilesReady();
  }
}

/**
 * Checks if both files are ready for processing
 */
function checkFilesReady() {
  uploadBtn.disabled = !(branchInfoFile && employeeDataFile);
}

/**
 * Processes both uploaded files
 */
function processFiles() {
  if (!branchInfoFile || !employeeDataFile) {
    showNotification("Please select both files first.", "error");
    return;
  }

  // Show loading indicator
  loadingIndicator.classList.remove("hidden");
  dataContainer.classList.add("hidden");

  // Reset data
  employeeData = [];
  branchManagers = {};
  branchMailIDs = {};

  // Process branch info file first
  processFile(branchInfoFile, true)
    .then(() => {
      // Then process employee data file
      return processFile(employeeDataFile, false);
    })
    .then(() => {
      // Group employees by branch and display
      displayEmployeesByBranch();

      // Hide loading and show data
      loadingIndicator.classList.add("hidden");
      dataContainer.classList.remove("hidden");

      showNotification("Data processed successfully!", "success");
    })
    .catch((error) => {
      loadingIndicator.classList.add("hidden");
      showNotification("Error processing files: " + error.message, "error");
    });
}

/**
 * Processes a file (either branch info or employee data)
 */
function processFile(file, isBranchInfo) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = function (e) {
      try {
        if (file.name.endsWith(".csv")) {
          Papa.parse(e.target.result, {
            header: true,
            skipEmptyLines: true,
            complete: function (results) {
              if (isBranchInfo) {
                processBranchInformation(results.data);
              } else {
                processEmployeeData([
                  Object.keys(results.data[0]), // Headers
                  ...results.data.map((row) => Object.values(row)), // Values
                ]);
              }
              resolve();
            },
            error: function (error) {
              reject(new Error("CSV parsing error: " + error.message));
            },
          });
        } else {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: "array" });
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];

          if (isBranchInfo) {
            const branchInfoData = XLSX.utils.sheet_to_json(worksheet);
            processBranchInformation(branchInfoData);
          } else {
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            processEmployeeData(jsonData);
          }
          resolve();
        }
      } catch (error) {
        reject(error);
      }
    };

    reader.onerror = function () {
      reject(new Error("Error reading file."));
    };

    if (file.name.endsWith(".csv")) {
      reader.readAsText(file);
    } else {
      reader.readAsArrayBuffer(file);
    }
  });
}

/**
 * Processes branch information from the Excel/CSV data
 */
function processBranchInformation(data) {
  if (!data || data.length === 0) {
    showNotification("Branch information file is empty or invalid.", "error");
    return;
  }

  // Expected column names (case insensitive)
  const branchColumnNames = ["Branch"];
  const managerColumnNames = ["Area Manager"];
  const mailColumnNames = ["Mail id"];

  data.forEach((row) => {
    let branch, areaManager, mailId;

    // Find the correct column names (case insensitive)
    const rowKeys = Array.isArray(row) ? [] : Object.keys(row);

    const branchKey = rowKeys.find((key) => branchColumnNames.includes(key));
    const managerKey = rowKeys.find((key) => managerColumnNames.includes(key));
    const mailKey = rowKeys.find((key) => mailColumnNames.includes(key));

    if (!Array.isArray(row)) {
      branch = branchKey ? row[branchKey] : null;
      areaManager = managerKey ? row[managerKey] : null;
      mailId = mailKey ? row[mailKey] : null;

      // Clean up values
      if (branch) {
        branch = branch.toString().trim();

        // Store the branch information
        if (areaManager) {
          branchManagers[branch] = areaManager.toString().trim();
        }
        if (mailId) {
          branchMailIDs[branch] = mailId.toString().trim();
        }
      }
    }
  });

  // Validate we got some branch information
  if (
    Object.keys(branchManagers).length === 0 &&
    Object.keys(branchMailIDs).length === 0
  ) {
    showNotification(
      "Warning: No branch manager or mail IDs found in the branch information file.",
      "warning"
    );
  }
}

// ... [rest of the existing functions remain the same, including processEmployeeData, displayEmployeesByBranch, prepareEmail, etc.] ...

/**
 * Processes the raw data from Excel/CSV for employee data
 */
/**
 * Processes the raw data from Excel/CSV for employee data
 */
/**
 * Processes the raw data from Excel/CSV for employee data
 */
function processEmployeeData(jsonData) {
  if (!jsonData || jsonData.length < 2) {
    showNotification("Invalid or empty data.", "error");
    return;
  }

  // Extract headers (first row)
  const headers = jsonData[0].map(h => h.toString().trim()); // Trim all headers

  // Map column indices (case insensitive)
  const columnIndices = {
    srNo: headers.findIndex(h => 
      h.match(/^(Sr_No|Sr No|S\.No)$/i)
    ),
    empId: headers.findIndex(h => 
      h.match(/^(Emp Id|Employee ID)$/i)
    ),
    empName: headers.findIndex(h => 
      h.match(/^(Emp Name|Employee Name)$/i)
    ),
    branch: headers.findIndex(h => 
      h.match(/^(BRANCH|Branch|Branch Name)$/i)
    ),
    department: headers.findIndex(h => 
      h.match(/^Department$/i)
    ),
    userId: headers.findIndex(h => 
      h.match(/^(Userid|User Id)$/i)
    ),
    userPassword: headers.findIndex(h => 
      h.match(/^User Password$/i)
    ),
    active: headers.findIndex(h => 
      h.match(/^Active$/i)
    ),
    mailId: headers.findIndex(h => 
      h.match(/^(Mail id|Branch Email)$/i)
    ),
    areaManager: headers.findIndex(h => 
      h.match(/^(Area Manager|Area Manager Email)$/i)
    ),
  };

  // Check if essential columns are found
  const essentialColumns = [
    "empId",
    "empName",
    "branch",
    "userId",
    "userPassword",
  ];
  const missingColumns = essentialColumns.filter(
    (col) => columnIndices[col] === -1
  );

  if (missingColumns.length > 0) {
    showNotification(`Missing required columns: ${missingColumns.join(", ")}`, "error");
    return;
  }

  // Reset employee data
  employeeData = [];

  // Process each row
  for (let i = 1; i < jsonData.length; i++) {
    const row = jsonData[i];
    if (!row || row.length === 0) continue;

    // Get branch name and trim it
    const branch = row[columnIndices.branch] 
      ? row[columnIndices.branch].toString().trim()
      : "";

    const employee = {
      srNo: row[columnIndices.srNo] || i,
      empId: row[columnIndices.empId] ? row[columnIndices.empId].toString().trim() : "",
      empName: row[columnIndices.empName] ? row[columnIndices.empName].toString().trim() : "",
      branch: branch,
      department: row[columnIndices.department] ? row[columnIndices.department].toString().trim() : "",
      userId: row[columnIndices.userId] ? row[columnIndices.userId].toString().trim() : "",
      userPassword: row[columnIndices.userPassword] ? row[columnIndices.userPassword].toString().trim() : "",
      active: row[columnIndices.active] ? row[columnIndices.active].toString().trim() : "",
    };

    // Only add if we have required fields
    if (employee.empId && employee.empName && employee.userId && employee.userPassword) {
      employeeData.push(employee);

      // Store branch info if available
      if (branch) {
        // Store mail ID if available
        if (columnIndices.mailId !== -1 && row[columnIndices.mailId] && !branchMailIDs[branch]) {
          branchMailIDs[branch] = row[columnIndices.mailId].toString().trim();
        }

        // Store area manager if available
        if (columnIndices.areaManager !== -1 && row[columnIndices.areaManager] && !branchManagers[branch]) {
          branchManagers[branch] = row[columnIndices.areaManager].toString().trim();
        }
      }
    }
  }

  // After processing all data, display it
  displayEmployeesByBranch();
}

/**
 * Prepares email for a specific branch with dynamic month/year
 */
function prepareEmail(branchName, employees) {
  if (!employees || employees.length === 0) {
    showNotification("No employees found for this branch.", "error");
    return;
  }

  // Hide data container and show email preview
  dataContainer.classList.add("hidden");
  emailPreview.classList.remove("hidden");

  // Get branch email and manager email
  const branchEmail = branchMailIDs[branchName] || "";
  const areaManagerEmail = branchManagers[branchName] || "";

  // Set recipients
  emailToElement.textContent = branchEmail;
  
  // Set CC list (including HR emails and area manager if available)
  const ccList = [...CC_LIST];
  if (areaManagerEmail) {
    ccList.push(areaManagerEmail);
  }
  emailCCElement.textContent = ccList.join(", ");

  // Get current month and year
  const now = new Date();
  const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", 
                     "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const currentMonth = monthNames[now.getMonth()];
  const currentYear = now.getFullYear().toString().slice(-2);

  // Create subject with dynamic month/year
  const employeeNames = employees.map(e => e.empName).join(", ");
  const subject = `ID Password New Joining ${branchName} Store ${currentMonth}-${currentYear} ll ${employeeNames}`;
  emailSubjectElement.textContent = subject;

  // Create HTML email body
  let htmlBody = `Dear Team,<br><br>`;
  htmlBody += `Please find below mentioned ID & Password. Kindly start mobile Punching from today onwards.<br><br>`;
  
  htmlBody += `<table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse; width: 100%;">`;
  htmlBody += `<thead><tr style="background-color: #f2f2f2;">`;
  htmlBody += `<th>Sr No</th><th>Emp ID</th><th>Name</th>`;
  htmlBody += `<th>Branch</th><th>User ID</th><th>Password</th>`;
  htmlBody += `</tr></thead><tbody>`;
  
  employees.forEach((emp, index) => {
    htmlBody += `<tr>`;
    htmlBody += `<td>${index + 1}</td>`;
    htmlBody += `<td>${emp.empId}</td>`;
    htmlBody += `<td>${emp.empName}</td>`;
    htmlBody += `<td>${emp.branch}</td>`;
    htmlBody += `<td>${emp.userId}</td>`;
    htmlBody += `<td>${emp.userPassword}</td>`;
    htmlBody += `</tr>`;
  });
  
  htmlBody += `</tbody></table><br><br>`;
  htmlBody += `Thank you,<br>HR Team`;

  // Create plain text version
  let plainBody = `Dear Team,\n\n`;
  plainBody += `Please find below mentioned ID & Password. Kindly start mobile Punching from today onwards.\n\n`;
  
  // Create table header
  plainBody += `+------+----------+----------------------+------------------+----------+----------+\n`;
  plainBody += `| SrNo | Emp ID   | Name                 | Branch           | User ID  | Password |\n`;
  plainBody += `+------+----------+----------------------+------------------+----------+----------+\n`;
  
  // Add rows
  employees.forEach((emp, index) => {
    plainBody += `| ${(index + 1).toString().padEnd(4)} `;
    plainBody += `| ${emp.empId.padEnd(8)} `;
    plainBody += `| ${emp.empName.padEnd(20)} `;
    plainBody += `| ${emp.branch.padEnd(16)} `;
    plainBody += `| ${emp.userId.padEnd(8)} `;
    plainBody += `| ${emp.userPassword.padEnd(8)} |\n`;
  });
  
  plainBody += `+------+----------+----------------------+------------------+----------+----------+\n\n`;
  plainBody += `Thank you,\nHR Team`;

  // Set the email content
  emailBodyElement.innerHTML = htmlBody;
  emailBodyElement.dataset.plainText = plainBody;

  // Store branch data for sending
  emailPreview.dataset.branchName = branchName;
  emailPreview.dataset.employees = JSON.stringify(employees);
}

/**
 * Prepares email for a specific branch
 */
/**
 * Prepares email for a specific branch
 */
/**
 * Prepares email for a specific branch
 */
/**
 * Prepares email for a specific branch
 */
function prepareEmail(branchName, employees) {
  if (!employees || employees.length === 0) {
    showNotification("No employees found for this branch.", "error");
    return;
  }

  // Hide data container and show email preview
  dataContainer.classList.add("hidden");
  emailPreview.classList.remove("hidden");

  // Get branch info from the branchMailIDs and branchManagers objects
  const branchEmail = branchMailIDs[branchName] || "";
  const areaManagerEmail = branchManagers[branchName] || "";

  // Set email recipient from branch Mail ID
  emailToElement.textContent = branchEmail;

  // Set CC list - Add area manager to CC list if available
  const ccList = [...CC_LIST];
  if (areaManagerEmail) {
    ccList.push(areaManagerEmail);
  }
  emailCCElement.textContent = ccList.join(", ");

  // Create employee name list for subject
  const employeeNames = employees.map((emp) => emp.empName).join(", ");

  // Get current month abbreviation and year
  const currentDate = new Date();
  const monthAbbr = currentDate.toLocaleString('default', { month: 'short' });
  const currentYear = currentDate.getFullYear().toString().slice(-2); // Last 2 digits of year

  // Set email subject with dynamic month and year
  const subject = `ID Password New Joining ${branchName} Store ${monthAbbr}-${currentYear} ll ${employeeNames}`;
  emailSubjectElement.textContent = subject;

  // Format email body - HTML version
  let bodyHTML = `Dear Team,<br><br>`;
  bodyHTML += `Please find below mentioned ID & Password. Kindly start mobile Punching from today onwards.<br><br>`;

  // Add employee table (HTML version)
  bodyHTML += `<table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse;">
        <thead>
            <tr style="background-color: #f2f2f2;">
                <th>Sr No</th>
                <th>Emp ID</th>
                <th>Name</th>
                <th>Branch</th>
                <th>User ID</th>
                <th>Password</th>
            </tr>
        </thead>
        <tbody>`;

  employees.forEach((employee, index) => {
    bodyHTML += `<tr>
            <td>${index + 1}</td>
            <td>${employee.empId}</td>
            <td>${employee.empName}</td>
            <td>${employee.branch}</td>
            <td>${employee.userId}</td>
            <td>${employee.userPassword}</td>
        </tr>`;
  });

  bodyHTML += `</tbody></table><br><br>`;
  bodyHTML += `Thank you,<br>HR Team`;

  // Format email body - Plain text version with borders
  let bodyText = `Dear Team,\n\n`;
  bodyText += `Please find below mentioned ID & Password. Kindly start mobile Punching from today onwards.\n\n`;

  // Calculate column widths
  const colWidths = [6, 10, 20, 15, 15, 15]; // Adjust these based on your data
  
  // Helper function to pad text to a specific width
  const padText = (text, width) => {
    return (text + ' '.repeat(width)).substring(0, width);
  };

  // Create table header with borders
  const headers = ['Sr No', 'Emp ID', 'Name', 'Branch', 'User ID', 'Password'];
  let headerLine = '';
  let separatorLine = '';
  
  headers.forEach((header, i) => {
    headerLine += `| ${padText(header, colWidths[i])} `;
    separatorLine += `+${'-'.repeat(colWidths[i] + 2)}`;
  });
  
  headerLine += '|';
  separatorLine += '+';
  
  bodyText += separatorLine + '\n';
  bodyText += headerLine + '\n';
  bodyText += separatorLine + '\n';

  // Add employee data with borders
  employees.forEach((employee, index) => {
    let rowLine = '';
    const rowData = [
      (index + 1).toString(),
      employee.empId,
      employee.empName,
      employee.branch,
      employee.userId,
      employee.userPassword
    ];
    
    rowData.forEach((data, i) => {
      rowLine += `| ${padText(data, colWidths[i])} `;
    });
    
    rowLine += '|';
    bodyText += rowLine + '\n';
  });

  bodyText += separatorLine + '\n\n';
  bodyText += `Thank you,\nHR Team`;

  // Store both versions in data attributes
  emailBodyElement.innerHTML = bodyHTML;
  emailBodyElement.dataset.plainText = bodyText;

  // Store current branch data for sending
  emailPreview.dataset.branchName = branchName;
  emailPreview.dataset.employees = JSON.stringify(employees);
  const emailActions = document.querySelector(".email-actions");
  emailActions.innerHTML =
    '<button id="backToListBtn" class="secondary-btn">Back to List</button>';

  // Store current branch data (though we may not need this anymore)
  emailPreview.dataset.branchName = branchName;
  emailPreview.dataset.employees = JSON.stringify(employees);
}

/**
 * Hides the email preview and shows the data container
 */
function hideEmailPreview() {
  emailPreview.classList.add("hidden");
  dataContainer.classList.remove("hidden");
}

/**
 * Simulates sending an email
 */
function sendEmail() {
  const branchName = emailPreview.dataset.branchName;

  // Here you would typically integrate with an email API
  // For this demo, we'll just show a success notification

  showNotification(
    `Email for ${branchName} branch has been sent successfully!`,
    "success"
  );

  // Hide email preview and show data container
  hideEmailPreview();
}

/**
 * Shows a notification message
 */
function showNotification(message, type = "info") {
  notificationMessage.textContent = message;

  // Set notification color based on type
  notification.className = "notification";
  notification.classList.add(`notification-${type}`);

  // Show notification
  notification.classList.remove("hidden");

  // Auto-hide after 5 seconds
  setTimeout(hideNotification, 5000);
}

/**
 * Hides the notification
 */
function hideNotification() {
  notification.classList.add("hidden");
}
