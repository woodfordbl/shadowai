/*
 * Copyright (c) Blake Woodford. All rights reserved. Licensed under the MIT license.
 */

// #region Imports
import { getUserProfile } from "../helpers/sso-helper";
import { filterUserProfileInfo } from "./../helpers/documentHelper";

let auth = null; // Auth token storage
let verifiedEmail = null; // Signup Email storage
let currentMSEmail = null; // Microsoft Email storage
//#endregion

// #region Initialize Office
Office.onReady((info) => {
  if (!Office.context.requirements.isSetSupported("ExcelApi", "1.7")) {
    // Check if Excel version is supported
    console.log("Sorry, this add-in only works with newer versions of Excel.");
  }
  if (info.host === Office.HostType.Excel) {
    getUserProfile(checkMSVerification); // Check if user is verified on startup
  }
});
// #endregion

//#region User Verifications

function updateEmail(id, email) {
  // TODO: Update the email address in the settings of the task pane
  var textarea = document.getElementById(id);
  textarea.placeholder = email;
  verifiedEmail = email; //updates global value
  const button = document.getElementById("auth-button");
  button.style.backgroundColor = "#a1a1a1";
  button.style.cursor = "default";
  console.log("Email updated.");
  return;
}

// Sends verification packet to server to verify on Microsoft Email || Input: Microsoft Email || Output: JSON Verification Packet
async function sendMSVerifyPacket(msEmail) {
  try {
    const response = await fetch("https://testing-f03s.onrender.com/verify-email", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ email: msEmail }),
    });

    if (response.ok) {
      const data = await response.json();
      return data;
    } else {
      console.error("Error:", response.status);
      return null;
    }
  } catch (error) {
    console.error("Error:", error);
    return null;
  }
}

// Checks verification of Micrsoft Email || Verified: logs auth, Not-Verified: does nothing
async function checkMSVerification(result) {
  const msEmail = filterUserProfileInfo(result)[0];

  if (msEmail === null) {
    console.log("No Microsoft Email");
    return;
  } else {
    currentMSEmail = msEmail;
  }

  const verificationStatus = await sendMSVerifyPacket(msEmail);

  if (!verificationStatus) {
    return;
  }

  if (verificationStatus.verification) {
    // On success, update authtoken and stored email
    auth = verificationStatus.authToken;
    updateEmail("signup-email", verificationStatus.stripeEmail);
  } else {
    // TODO: Handle unverified user
    console.log("User not verified on MS Email");
    return;
  }
}

async function verifyUser(signUpEmail) {
  try {
    const response = await fetch("https://testing-f03s.onrender.com/verify-user", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ email: signUpEmail, msEmail: currentMSEmail }),
    });

    if (response.ok) {
      const data = await response.json();
      return data;
    } else {
      console.error("Error:", response.status);
    }
  } catch (error) {
    console.error("Error:", error);
  }
}

async function authenticateUser(emailId, statusId) {
  if (!(verifiedEmail === null)) {
    console.log("User already authenticated");
    return;
  }
  const signUpEmail = document.getElementById(emailId).value;
  const status = document.getElementById(statusId);

  if (signUpEmail === "") {
    status.value = "Please enter an email address.";
    return;
  }

  try {
    const verificationPacket = await verifyUser(signUpEmail);

    if (verificationPacket.verification) {
      // On success, update authtoken and stored email
      auth = verificationPacket.authToken;
      updateEmail(emailId, signUpEmail);
      status.value = verificationPacket.verificationMessage;
      return;
    } else {
      status.value = verificationPacket.verificationMessage;
      return;
    }
  } catch (error) {
    console.error("Error:", error);
    status.value = "Unable to verify email, please try again.";
    return;
  }
}

async function sendDeAuth(emailId) {
  if (verifiedEmail === null || currentMSEmail === undefined) {
    console.log("No current user");
    return { message: "No current user" };
  }

  try {
    const response = await fetch("https://testing-f03s.onrender.com/deauth-user", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ email: verifiedEmail, msEmail: currentMSEmail, authToken: auth }),
    });

    if (response.ok) {
      const data = await response.json();
      auth = null;
      verifiedEmail = null;
      document.getElementById(emailId).value = "";
      document.getElementById(emailId).placeholder = "Email";
      return data;
    } else {
      console.error("Error:", response.status);
      return { message: "Unable to deauthenticate user, please try again." };
    }
  } catch (error) {
    console.error("Error:", error);
    return { message: "Unable to deauthenticate user, please try again." };
  }
}

async function deAuthenticateUser(emailId, statusId) {
  const deauthpacket = await sendDeAuth(emailId);
  const status = document.getElementById(statusId);
  status = deauthpacket.message;
}

//#endregion

//#region Handle switching between tabs
document.addEventListener("DOMContentLoaded", function () {
  "use strict";

  // Store references to the tabs and tab contents
  var tabs = document.querySelectorAll(".tab");
  var tabContents = document.querySelectorAll(".tab-content");

  // Function to switch tabs
  function switchTab(tabId) {
    // Hide all tab contents and deactivate all tabs
    tabContents.forEach(function (tabContent) {
      tabContent.style.display = "none";
    });
    tabs.forEach(function (tab) {
      tab.classList.remove("active");
    });

    // Show the selected tab content and activate the corresponding tab
    var selectedTabContent = document.getElementById(tabId);
    var selectedTab = document.querySelector('a[data-tab="' + tabId + '"]');
    selectedTabContent.style.display = "flex";
    selectedTab.classList.add("active");
  }

  // Add click event listeners to the tabs
  tabs.forEach(function (tab) {
    tab.addEventListener("click", function () {
      var tabId = this.getAttribute("data-tab");
      switchTab(tabId);
    });
  });

  // Set the initial active tab
  switchTab("TabCreate");
});

//#endregion

//#region Functions to load and display text

// Loader element to display elipsis while waiting for GPT response
let loadInterval = null;

function loader(element) {
  element.textContent = "";

  loadInterval = setInterval(function () {
    // Update the text content of the loading indicator
    element.textContent += ". ";

    // If the loading indicator has reached three dots, reset it
    if (element.textContent === ". . . . ") {
      element.textContent = "";
    }
  }, 300);
}

function typeText(element, wrappedText, uniqueId, type) {
  let index = 0;
  let currentSpan = null;

  if (wrappedText === undefined || wrappedText === "") {
    return;
  }

  let interval = setInterval(function () {
    if (index < wrappedText.length) {
      const char = wrappedText.charAt(index);

      if (char === "<" && wrappedText.charAt(index + 1) !== "/") {
        const closingBracketIndex = wrappedText.indexOf(">", index);
        if (closingBracketIndex !== -1) {
          const spanTag = wrappedText.substring(index, closingBracketIndex + 1);
          element.innerHTML += spanTag;
          currentSpan = element.lastElementChild;
          index = closingBracketIndex;
        }
      } else if (char === "<" && wrappedText.charAt(index + 1) === "/") {
        const closingBracketIndex = wrappedText.indexOf(">", index);
        if (closingBracketIndex !== -1) {
          currentSpan = null;
          index = closingBracketIndex;
        }
      } else if (currentSpan !== null) {
        currentSpan.innerHTML += char;
      } else {
        element.innerHTML += char;
      }

      index++;
    } else {
      clearInterval(interval);
    }
  }, 20);

  if (uniqueId !== undefined) {
    document.getElementById(`insert-${uniqueId}`).addEventListener("click", () => {
      insertFormula(`insert-${uniqueId}`);
    });
    document.getElementById(`retry-${uniqueId}`).addEventListener("click", () => {
      retryQuery(uniqueId, type);
    });
  }
}

function generateUniqueId() {
  const timestamp = Date.now();
  const randomNumber = Math.random();
  const hexadecimalString = randomNumber.toString(16);

  return `${timestamp}-${hexadecimalString}`;
}

function outputText(isAi, value, uniqueId) {
  return `
        <div class="wrapper ${isAi && "ai"}">
            <div class="chat">
                <div class="profile">${isAi ? "🥷" : "👤"}</div>
                <div class="message ${isAi && "ai-message"}" id=${isAi ? "output-" : "input-"}${uniqueId}>${value}</div>
                ${
                  isAi
                    ? `<button id="insert-${uniqueId}" title="Paste into Current Cell" class="insert"><i class="fas fa-clipboard"></i></button>`
                    : `<button id="retry-${uniqueId}" title="Resend" class="insert"><i class="fas fa-redo"></i></button>`
                }
            </div>
        </div>
        `;
}
//#endregion

//#region Functions to handle insert/retry button
async function retryQuery(elementId, type) {
  if (elementId === undefined) {
    return;
  }
  if (type === undefined) {
    return;
  }
  const inputElement = document.getElementById("input-"+elementId);
  const outputElement = document.getElementById("output-"+elementId);

  // Clear output div
  outputElement.innerHTML = "";

  // Initiate Loader
  loader(outputElement);

  let gptResponse = "";

  // Get new GPT text
  if (type==="formula") {
  gptResponse = (await getFormula(inputElement.textContent.trim())).trim(); // Get GPT text
  } else if (type==="explain") {
    gptResponse = (await getExplain(inputElement.textContent.trim())).trim(); // Get GPT text
  } else {
    return;
  }

  // Check if '=' exists in the gptResponse
  const equalIndex = gptResponse.indexOf("=");

  if (equalIndex !== -1) {
    // Remove text to the left of '='
    const evaluatedText = gptResponse.slice(equalIndex).trim();

    clearInterval(loadInterval);
    outputElement.innerHTML = "";

    // Create HTML element from GPT text
    typeText(outputElement, wrapCellReferencesWithSpan(evaluatedText));
  } else {
    clearInterval(loadInterval);
    outputElement.innerHTML = "";

    // Create HTML element from GPT text
    typeText(outputElement, wrapCellReferencesWithSpan(gptResponse));
  }
}

function copyToClipboard(text) {
  // Create a temporary textarea element
  const textarea = document.createElement("textarea");
  textarea.value = text;

  // Append the textarea to the document body
  document.body.appendChild(textarea);

  // Select the text within the textarea
  textarea.select();

  // Copy the selected text to the clipboard
  document.execCommand("copy");

  // Remove the temporary textarea
  document.body.removeChild(textarea);
}

function insertFormula(buttonId) {
  const messageId = buttonId.replace("insert-", "output-");
  const messageElement = document.getElementById(messageId);
  copyToClipboard(messageElement.textContent.trim()); //copy to user clipboard just in case

  Excel.run(function (context) {
    // Load the selected range and its properties
    var range = context.workbook.getSelectedRange();
    range.load("rowCount");
    range.load("columnCount");

    return context
      .sync()
      .then(function () {
        // Check if only one cell is selected
        if (range.rowCount === 1 && range.columnCount === 1) {
          // Set the formula in the current cell
          range.formulas = [[messageElement.textContent.trim()]];
        }
      })
      .then(context.sync);
  }).catch(function (error) {
    console.log(error);
  });
}

function copyElementToClipboard(elementId) {
  const element = document.getElementById(elementId);
  copyToClipboard(element.textContent.trim());
}
//#endregion

//#region Handle prompt inputs and buttons
function autoResizeTextarea(textareaIds) {
  textareaIds.forEach(function (textareaId) {
    var textarea = document.getElementById(textareaId);

    var minHeight = getComputedStyle(textarea).height.replace(/px$/, "");
    var maxHeight = getComputedStyle(textarea).maxHeight.replace(/px$/, "");

    textarea.addEventListener("input", function (event) {
      // Calculate the scroll height of the textarea content
      textarea.style.height = "1px";
      var scrollHeight = textarea.scrollHeight;

      // Check if the scroll height exceeds the maximum height
      if (scrollHeight > maxHeight) {
        textarea.style.overflowY = "scroll"; // Display the scrollbar
        textarea.style.height = maxHeight + "px"; // Set the height to the maximum height
      } else if (scrollHeight < minHeight) {
        textarea.style.overflowY = "hidden"; // Hide the scrollbar
        textarea.style.height = minHeight + "px"; // Set the height to fit the content
      } else {
        textarea.style.overflowY = "hidden"; // Hide the scrollbar
        textarea.style.height = scrollHeight + "px"; // Set the height to fit the content
      }
    });
  });
}

// Format textarea elements
document.addEventListener("DOMContentLoaded", function () {
  autoResizeTextarea(["generateInput", "explainInput", "vbaInput"]); // Pass an array of textarea IDs
});

function submitOnEnter(elementIds) {
  elementIds.forEach(function (id) {
    var element = document.getElementById(id);
    if (element) {
      element.addEventListener("keydown", function (event) {
        if (event.key === "Enter" && !event.shiftKey) {
          event.preventDefault();
          if (id === "generateInput") {
            createFormula(id);
          } else if (id === "explainInput") {
            explainFormula(id);
          } else if (id === "vbaInput") {
            vbaFormula(id);
          }
        }
      });
    }
  });
}

var inputIds = ["generateInput", "explainInput", "vbaInput"]; // Provide the element IDs as an array
submitOnEnter(inputIds);

function getCurrent() {
  Excel.run(function (context) {
    var selectedRange = context.workbook.getSelectedRange();
    selectedRange.load("formulas");

    return context.sync().then(function () {
      if (selectedRange.formulas.length === 1) {
        var formula = selectedRange.formulas[0][0];
        var formulaElement = document.getElementById("explainInput");
        formulaElement.textContent = formula;
      }
    });
  }).catch(function (error) {
    console.error(error);
  });
}

let generateHTML, explainHTML, vbaHTML;
document.addEventListener("DOMContentLoaded", (event) => {
  generateHTML = document.getElementById("output").innerHTML;
  explainHTML = document.getElementById("explain-output").innerHTML;
  vbaHTML = document.getElementById("code-output").innerHTML;
});

function refreshPage(id, htmlTag) {
  output = document.getElementById(id);
  output.innerHTML = htmlTag;
  if (htmlTag === explainHTML) {
    initializeExplainButtons();
  }
  if (htmlTag === generateHTML) {
    initializeGenerateButtons();
  }
}

//#endregion

//#region Functions to handle server querys || TODO: Move to separate file

// Process user input from textarea elements
function processInput(id) {
  // Get text from div
  var textInput = document.getElementById(id);
  if (textInput.tagName === "TEXTAREA") {
    var text = textInput.value.trim();
  } else if (textInput.tagName === "DIV") {
    var text = textInput.textContent.trim();
  } else {
    return;
  }
  if (text === "") {
    return;
  } else {
    // Clear the form after text entry
    textInput.value = "";
    return text;
  }
}

// Create a formula from user input
async function getFormula(query) {
  try {
    const response = await fetch("https://testing-f03s.onrender.com/formula", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ prompt: query, token: auth }),
    });

    if (response.ok) {
      const data = await response.json();
      return data.bot;
    } else {
      console.error("Error:", response.status);
    }
  } catch (error) {
    console.error("Error:", error);
  }
}
// Explain a formula from user input
async function getExplain(query) {
  try {
    const response = await fetch("https://testing-f03s.onrender.com/explain", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ prompt: query, token: auth }),
    });

    if (response.ok) {
      const data = await response.json();
      return data.bot;
    } else {
      console.error("Error:", response.status);
    }
  } catch (error) {
    console.error("Error:", error);
  }
}
// Get VBA file from user input
async function getVBA(query) {
  try {
    const response = await fetch("https://testing-f03s.onrender.com/vba", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ prompt: query, token: auth }),
    });

    if (response.ok) {
      const data = await response.json();
      return data.bot;
    } else {
      console.error("Error:", response.status);
    }
  } catch (error) {
    console.error("Error:", error);
  }
}

async function createFormula(id) {
  const text = processInput(id);

  // Clear output div
  var divElement = document.getElementById("output");
  divElement.innerHTML = "";

  var uniqueId = generateUniqueId();

  const wrappedText = wrapCellReferencesWithSpan(text);

  // Create HTML element from user text
  var outputElement = document.getElementById("output");
  var output = outputText(false, wrappedText, uniqueId);
  outputElement.innerHTML += output;

  // Pre-create HTML element for GPT text
  outputElement.innerHTML += outputText(true, "", uniqueId);

  // Get the pre-created HTML element and add loader element
  const messageDiv = document.getElementById("output-" + uniqueId);
  loader(messageDiv);
  const gptResponse = (await getFormula(text)).trim(); // Get GPT text

  // Check if '=' exists in the gptResponse
  const equalIndex = gptResponse.indexOf("=");

  if (equalIndex !== -1) {
    // Remove text to the left of '='
    const evaluatedText = gptResponse.slice(equalIndex).trim();

    clearInterval(loadInterval);
    messageDiv.innerHTML = "";

    // Create HTML element from GPT text
    typeText(messageDiv, wrapCellReferencesWithSpan(evaluatedText), uniqueId, "formula");
  } else {
    clearInterval(loadInterval);
    messageDiv.innerHTML = "";

    // Create HTML element from GPT text
    typeText(messageDiv, wrapCellReferencesWithSpan(gptResponse), uniqueId, "formula");
  }
}
async function explainFormula(id) {
  const text = processInput(id);

  // Get text from div
  var textInput = document.getElementById(id);

  // Clear output div
  var divElement = document.getElementById("explain-output");
  divElement.innerHTML = "";

  var uniqueId = generateUniqueId();
  const wrappedText = wrapCellReferencesWithSpan(text);
  // Create HTML element from user text
  var outputElement = document.getElementById("explain-output");
  var output = outputText(false, wrappedText, uniqueId);
  outputElement.innerHTML += output;

  // Clear the form after text entry
  textInput.textContent = "";

  // Pre-create HTML element for GPT text
  outputElement.innerHTML += outputText(true, "", uniqueId);

  // Get the pre-created HTML element and add loader element
  const messageDiv = document.getElementById("output-" + uniqueId);
  loader(messageDiv);
  const gptResponse = (await getExplain(text)).trim(); // Get GPT text

  // Check if '=' exists in the gptResponse
  const equalIndex = gptResponse.indexOf("=");

  if (equalIndex !== -1) {
    // Remove text to the left of '='
    const evaluatedText = gptResponse.slice(equalIndex).trim();

    clearInterval(loadInterval);
    messageDiv.innerHTML = "";

    // Create HTML element from GPT text
    typeText(messageDiv, wrapCellReferencesWithSpan(evaluatedText), uniqueId, "explain");
  } else {
    clearInterval(loadInterval);
    messageDiv.innerHTML = "";

    // Create HTML element from GPT text
    typeText(messageDiv, wrapCellReferencesWithSpan(gptResponse), uniqueId, "explain");
  }
}
async function vbaFormula(id) {
  const text = processInput(id);

  // Output div
  var outputElement = document.getElementById("code-output");

  // Clear output div and text input
  outputElement.innerHTML = "";

  loader(outputElement);

  const gptResponse = await getVBA(text);

  var regex = /```([\s\S]*?)```/;
  var match = gptResponse.match(regex);

  if (match) {
    var result = match[1]; // Return the captured group
  } else {
    var result = gptResponse;
  }

  clearInterval(loadInterval);
  outputElement.innerHTML = "";

  result = formatCodeWithHLJS(result);

  typeText(outputElement, result);
}
//#endregion

//#region Functions to handle text and code formatting
function wrapCellReferencesWithSpan(text) {
  const cellReferenceRegex = /\$?[A-Za-z]+\$?\d+(?::\$?[A-Za-z]+\$?\d+)?/g;
  const cellReferences = text.match(cellReferenceRegex);

  const colors = ["#00A7E1", "#4094DE", "#6A7ED1", "#8A65BA", "#9F4999", "#A82B70"]; // List of colors
  const colorMap = {}; // Map to track assigned colors

  const wrappedText = text.replace(cellReferenceRegex, function (match) {
    if (!colorMap[match]) {
      const colorIndex = Object.keys(colorMap).length % colors.length;
      colorMap[match] = colors[colorIndex];
    }
    const spanStyle = `color: ${colorMap[match]}`;
    return `<span class="cell-reference" style="${spanStyle}">${match}</span>`;
  });
  return wrappedText;
}

var hoveredElement = null;
function handleHover(event) {
  hoveredElement = event.target;
}

// Function to toggle the cell reference format
function toggleCellReferenceFormat() {
  if (hoveredElement) {
    var currentText = hoveredElement.textContent;
    var newText = "";

    // Regular expressions for pattern matching
    var rangePattern = /^([$]?[A-Z]+)([$]?[0-9]+):([$]?[A-Z]+)([$]?[0-9]+)$/;
    var cellPattern = /^([$]?[A-Z]+)([$]?[0-9]+)$/;

    // Check if it's a range
    if (rangePattern.test(currentText)) {
      newText = currentText.replace(rangePattern, function (match, startCol, startRow, endCol, endRow) {
        if (startRow.startsWith("$")) {
          if (!endCol.startsWith("$")) {
            // Remove '$' at the start if present
            startRow = startRow.substr(1);
            endRow = endRow.substr(1);
          }
        } else {
          if (!endCol.startsWith("$")) {
            // Add '$' at the start
            startRow = "$" + startRow;
            endRow = "$" + endRow;
          }
        }
        if (startCol.startsWith("$")) {
          // Remove '$' at the start if present
          startCol = startCol.substr(1);
          endCol = endCol.substr(1);
        } else {
          // Add '$' at the start
          startCol = "$" + startCol;
          endCol = "$" + endCol;
        }

        return startCol + startRow + ":" + endCol + endRow;
      });
    }
    // Check if it's a single cell
    else if (cellPattern.test(currentText)) {
      newText = currentText.replace(cellPattern, function (match, col, row) {
        if (row.startsWith("$")) {
          if (!col.startsWith("$")) {
            // Remove '$' at the start if present
            row = row.substr(1);
          }
        } else {
          if (!col.startsWith("$")) {
            // Add '$' at the start
            row = "$" + row;
          }
        }
        if (col.startsWith("$")) {
          // Remove '$' at the start if present
          col = col.substr(1);
        } else {
          // Add '$' at the start
          col = "$" + col;
        }

        return col + row;
      });
    }

    // Update the text of the hovered element
    hoveredElement.textContent = newText;
  }
}

// Function to handle the F4 key press event
function handleKeyPress(event) {
  // Check if the pressed key is F4 (keyCode 115) and hovered element is a cell reference
  if (event.keyCode === 115 && hoveredElement && hoveredElement.classList.contains("cell-reference")) {
    toggleCellReferenceFormat();
  }
}
// Register the event listeners
document.addEventListener("mouseover", handleHover);
document.addEventListener("keydown", handleKeyPress);

function formatCodeWithHLJS(code) {
  var tempElement = document.createElement("div");
  tempElement.innerHTML = code;

  hljs.highlightBlock(tempElement);

  return tempElement.innerHTML;
}

// Handle code formatting on the page
function formatCode() {
  document.addEventListener("DOMContentLoaded", (event) => {
    hljs.highlightAll();
  });
}

formatCode();
//#endregion

//#region Button Clicks || TODO: Debug each button

// –––––––––– Example buttons –––––––––––

function initializeGenerateButtons() {
  document.getElementById("button-rec1").addEventListener("click", () => {
    // Create: Example
    createFormula("rec1");
  });
  document.getElementById("button-rec2").addEventListener("click", () => {
    // Create: Example
    createFormula("rec2");
  });
  document.getElementById("button-rec3").addEventListener("click", () => {
    // Create: Example
    createFormula("rec3");
  });
}
initializeGenerateButtons();

function initializeExplainButtons() {
  document.getElementById("explain1").addEventListener("click", () => {
    // Explain: Example
    explainFormula("explain1");
  });
  document.getElementById("explain2").addEventListener("click", () => {
    // Explain: Example
    explainFormula("explain2");
  });
  document.getElementById("explain3").addEventListener("click", () => {
    // Explain: Example
    explainFormula("explain3");
  });
  document.getElementById("explain4").addEventListener("click", () => {
    // Explain: Example
    explainFormula("explain4");
  });
  document.getElementById("explain5").addEventListener("click", () => {
    // Explain: Example
    explainFormula("explain5");
  });
  document.getElementById("explain6").addEventListener("click", () => {
    // Explain: Example
    explainFormula("explain6");
  });
}
initializeExplainButtons();

// –––––––––– Input buttons –––––––––––
document.getElementById("generateInput-button").addEventListener("click", () => {
  // Create formula button
  createFormula("generateInput");
});
document.getElementById("button-explainInput").addEventListener("click", () => {
  // Explain formula button
  explainFormula("explainInput");
});
document.getElementById("button-vbaInput").addEventListener("click", () => {
  // VBA formula button
  vbaFormula("vbaInput");
});

document.getElementById("generate-refresh").addEventListener("click", () => {
  // Generate: Refresh App
  refreshPage("output", generateHTML);
});
document.getElementById("explain-refresh").addEventListener("click", () => {
  // Explain: Refresh App
  refreshPage("explain-output", explainHTML);
});
document.getElementById("vba-refresh").addEventListener("click", () => {
  // VBA: Refresh App
  refreshPage("code-output", vbaHTML);
});

document.getElementById("explain-current").addEventListener("click", () => {
  // Explain: Get Current
  getCurrent();
});
document.getElementById("vba-current").addEventListener("click", () => {
  // VBA: Get Current
  getCurrent();
});

// –––––––––– Output buttons –––––––––––
document.getElementById("vba-copy").addEventListener("click", () => {
  // VBA: Copy to Clipboard
  copyElementToClipboard("code-output");
});

// –––––––––– Authentication buttons –––––––––––
const emailTextarea = document.getElementById("signup-email");
emailTextarea.addEventListener("focus", () => {
  const placeholder = emailTextarea.placeholder;
  if (emailTextarea.value === "") {
    emailTextarea.value = placeholder;
  }
});

emailTextarea.addEventListener("blur", () => {
  const placeholder = emailTextarea.placeholder;
  if (emailTextarea.value === placeholder) {
    emailTextarea.value = "";
  }
});

document.getElementById("auth-button").addEventListener("click", () => {
  authenticateUser("signup-email", "auth-status");
});
document.getElementById("deauth-button").addEventListener("click", () => {
  deAuthenticateUser("signup-email", "auth-status");
});
//#endregion
