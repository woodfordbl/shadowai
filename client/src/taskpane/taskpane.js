//#region Initializing Functions
Office.onReady().then(function () {
  if (!Office.context.requirements.isSetSupported("ExcelApi", "1.7")) {
    console.log("Sorry, this add-in only works with newer versions of Excel.");
  }
});
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

function typeText(element, wrappedText) {
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
                    ? `<button id="insert-${uniqueId}" title="Paste into Current Cell" class="insert" onclick="insertFormula(this.id)"><i class="fas fa-clipboard"></i></button>`
                    : `<button id="insert-${uniqueId}" title="Resend" class="insert" onclick="retryQuery(this.id)"><i class="fas fa-redo"></i></button>`
                }
            </div>
        </div>
        `;
}
//#endregion

//#region Functions to handle insert/retry button
async function retryQuery(elementId) {
  inputElement = document.getElementById(elementId.replace("insert-", "input-"));
  outputElement = document.getElementById("output-" + elementId.replace("insert-", ""));

  // Clear output div
  outputElement.innerHTML = "";

  // Initiate Loader
  loader(outputElement);

  // Get new GPT text
  const gptResponse = (await getGPT(inputElement.textContent.trim())).trim(); // Get GPT text

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
function autoResizeTextarea(textareaId) {
  textareaElement = document.getElementById(textareaId);

  minHeight = getComputedStyle(textareaElement).minHeight;
  maxHeight = getComputedStyle(textareaElement).maxHeight;

  textareaElement.style.height = minHeight; // Set the initial height to the minimum height

  textareaElement.addEventListener("input", function () {
    // Calculate the scroll height of the textareaElement content
    const height = textareaElement.scrollHeight; // Take exisiting height
    textareaElement.style.height = "1px"; // Set the height to 1px
    const scrollHeight = textareaElement.scrollHeight; // Measure scroll height
    textareaElement.style.height = `${height}px`; // Set the height back to the original height

    // Check if the scroll height exceeds the maximum height
    if (scrollHeight > maxHeight) {
      textareaElement.style.overflowY = "scroll"; // Display the scrollbar
      textareaElement.style.height = maxHeight; // Set the height to the maximum height
    } else if (scrollHeight < minHeight) {
      textareaElement.style.overflowY = "hidden"; // Hide the scrollbar
      textareaElement.style.height = minHeight; // Set the height to fit the content
    } else {
      textareaElement.style.height = "1px";
      textareaElement.style.height = textareaElement.scrollHeight + "px";
      textareaElement.style.overflowY = "hidden"; // Hide the scrollbar
    }
  });
}

// Format textarea elements
document.addEventListener("DOMContentLoaded", function () {
autoResizeTextarea("generateInput");
autoResizeTextarea("explainInput");
autoResizeTextarea("vbaInput");

});

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

function refreshApp() {
  location.reload();
}

function refreshPage(id) {
  output = document.getElementById(id);
  output.innerHTML = "";
}

//#endregion

//#region Functions to handle GPT querys

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
  console.log(text);
  if (text === "") {
    return;
  } else {
    // Clear the form after text entry
    textInput.value = "";
    return text;
}
}

async function getFormula(query) {
  try {
    const response = await fetch("https://testing-f03s.onrender.com/formula", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ prompt: query }),
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
async function getExplain(query) {
  try {
    const response = await fetch("https://testing-f03s.onrender.com/explain", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ prompt: query }),
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
async function getVBA(query) {
  try {
    const response = await fetch("https://testing-f03s.onrender.com/vba", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ prompt: query }),
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

async function createFormula(event, id) {
  event.preventDefault(); // Prevent form from submitting
  text = processInput(id);

  // Clear output div
  var divElement = document.getElementById("output");
  divElement.innerHTML = "";

  var uniqueId = generateUniqueId();

  wrappedText = wrapCellReferencesWithSpan(text);

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
    typeText(messageDiv, wrapCellReferencesWithSpan(evaluatedText));
  } else {
    clearInterval(loadInterval);
    messageDiv.innerHTML = "";

    // Create HTML element from GPT text
    typeText(messageDiv, wrapCellReferencesWithSpan(gptResponse));
  }
}
async function explainFormula(event, id) {
  event.preventDefault(); // Prevent form from submitting
  text = processInput(id);

  // Get text from div
  var textInput = document.getElementById(id);


  // Clear output div
  var divElement = document.getElementById("explain-output");
  divElement.innerHTML = "";

  var uniqueId = generateUniqueId();
  wrappedText = wrapCellReferencesWithSpan(text);
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
    typeText(messageDiv, wrapCellReferencesWithSpan(evaluatedText));
  } else {
    clearInterval(loadInterval);
    messageDiv.innerHTML = "";

    // Create HTML element from GPT text
    typeText(messageDiv, wrapCellReferencesWithSpan(gptResponse));
  }
}
async function vbaFormula(event, id) {
  event.preventDefault(); // Prevent form from submitting

  text = processInput(id);

  // Output div
  var outputElement = document.getElementById("code-output");

  // Clear output div and text input
  outputElement.innerHTML = "";

  loader(outputElement);

  gptResponse = await getVBA(text);

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
