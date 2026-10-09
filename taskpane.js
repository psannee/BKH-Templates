/* global document, Office, console, fetch */

// ==========================================
// CONFIGURATION
// ==========================================
// Default export URL for your published Google Sheet:
const GOOGLE_SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/1Ag3tpxxKgaRVPX4xUSVE35iN9BpPYNGr2saQk32-sAQ/export?format=csv";

// Store loaded templates in memory
let templatesDatabase = [];

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    initUI();
    loadTemplates();
  }
});

/**
 * Initializes static UI event listeners
 */
function initUI() {
  const refreshBtn = document.getElementById("refresh-templates-btn");
  if (refreshBtn) {
    refreshBtn.onclick = () => {
      showStatus("Refreshing templates...", "info");
      loadTemplates();
    };
  }
}

/**
 * Displays user notification banners in the taskpane
 */
function showStatus(message, type = "info") {
  const banner = document.getElementById("status-banner");
  if (!banner) return;
  
  banner.textContent = message;
  banner.className = `status-banner status-${type}`;
  banner.style.display = "block";

  if (type === "success" || type === "info") {
    setTimeout(() => {
      banner.style.display = "none";
    }, 4500);
  }
}

/**
 * Fetches the published Google Sheet CSV and generates the dynamic interface
 */
async function loadTemplates() {
  try {
    showStatus("Loading templates from database...", "info");
    const response = await fetch(GOOGLE_SHEET_CSV_URL, { cache: "no-cache" });
    
    if (!response.ok) {
      throw new Error(`HTTP error ${response.status}`);
    }

    const csvText = await response.text();
    templatesDatabase = parseTemplatesCSV(csvText);

    if (templatesDatabase.length === 0) {
      showStatus("No templates found in spreadsheet.", "warning");
      return;
    }

    renderWorkflowUI(templatesDatabase);
    showStatus(`Loaded ${templatesDatabase.length} workflows successfully.`, "success");
  } catch (error) {
    console.error("Failed to load templates:", error);
    showStatus("Could not load templates from Google Sheet. Check published URL.", "error");
  }
}

/**
 * Parses CSV text, handling multiline cells and quotes while preserving internal spacing
 */
function parseTemplatesCSV(csvText) {
  const rows = [];
  let currentRow = [];
  let currentVal = "";
  let inQuotes = false;

  for (let i = 0; i < csvText.length; i++) {
    const char = csvText[i];
    const nextChar = csvText[i + 1];

    if (char === '"') {
      if (inQuotes && nextChar === '"') {
        currentVal += '"';
        i++; // skip escaped quote
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === "," && !inQuotes) {
      currentRow.push(currentVal.trim());
      currentVal = "";
    } else if ((char === "\r" || char === "\n") && !inQuotes) {
      if (char === "\r" && nextChar === "\n") {
        i++;
      }
      currentRow.push(currentVal.trim());
      currentVal = "";
      if (currentRow.some((val) => val.length > 0)) {
        rows.push(currentRow);
      }
      currentRow = [];
    } else {
      currentVal += char;
    }
  }

  if (currentVal || currentRow.length > 0) {
    currentRow.push(currentVal.trim());
    if (currentRow.some((val) => val.length > 0)) {
      rows.push(currentRow);
    }
  }

  if (rows.length < 2) return [];

  // Identify column headers
  const headers = rows[0].map((h) => h.toLowerCase());
  const folderIdx = headers.findIndex((h) => h.includes("folder"));
  const workflowIdx = headers.findIndex((h) => h.includes("workflow") || h.includes("name"));
  const templateIdx = headers.findIndex((h) => h.includes("template") || h.includes("text"));
  const attachmentIdx = headers.findIndex((h) => h.includes("attachment") || h.includes("path"));

  const workflows = [];

  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row || row.length === 0) continue;

    const folder = folderIdx !== -1 && row[folderIdx] ? row[folderIdx] : "Universal";
    let workflowName = workflowIdx !== -1 && row[workflowIdx] ? row[workflowIdx] : "";
    let templateText = templateIdx !== -1 && row[templateIdx] ? row[templateIdx] : "";
    const attachmentPath = attachmentIdx !== -1 && row[attachmentIdx] ? row[attachmentIdx] : "";

    // Fallback if workflow name or template column layout varies
    if (!workflowName && templateText) {
      workflowName = folder !== "Universal" ? folder : "Send Response";
    } else if (!templateText && workflowName) {
      templateText = workflowName;
      workflowName = folder;
    }

    if (templateText) {
      workflows.push({
        folder: folder,
        name: workflowName,
        template: templateText,
        attachmentPath: attachmentPath
      });
    }
  }

  return workflows;
}

/**
 * Dynamically builds the accordion folders and workflow buttons
 */
function renderWorkflowUI(workflows) {
  const container = document.getElementById("workflows-container");
  if (!container) return;

  container.innerHTML = "";

  // Group workflows by folder
  const grouped = {};
  workflows.forEach((wf) => {
    const folderKey = wf.folder || "Universal";
    if (!grouped[folderKey]) {
      grouped[folderKey] = [];
    }
    grouped[folderKey].push(wf);
  });

  // 1. Render Universal / Top-Level Workflows first (if any)
  if (grouped["Universal"]) {
    const universalSec = document.createElement("div");
    universalSec.className = "universal-section";
    grouped["Universal"].forEach((wf) => {
      const btn = createWorkflowButton(wf);
      universalSec.appendChild(btn);
    });
    container.appendChild(universalSec);
    delete grouped["Universal"];
  }

  // 2. Render Accordion Folders
  Object.keys(grouped).forEach((folderName) => {
    const folderWrapper = document.createElement("div");
    folderWrapper.className = "accordion-group";

    const header = document.createElement("button");
    header.className = "accordion-header";
    header.innerHTML = `<span>📁 ${folderName}</span><span class="chevron">▾</span>`;

    const content = document.createElement("div");
    content.className = "accordion-content open";

    grouped[folderName].forEach((wf) => {
      const btn = createWorkflowButton(wf);
      content.appendChild(btn);
    });

    // Toggle accordion expansion
    header.onclick = () => {
      const isOpen = content.classList.contains("open");
      if (isOpen) {
        content.classList.remove("open");
        header.querySelector(".chevron").textContent = "▸";
      } else {
        content.classList.add("open");
        header.querySelector(".chevron").textContent = "▾";
      }
    };

    folderWrapper.appendChild(header);
    folderWrapper.appendChild(content);
    container.appendChild(folderWrapper);
  });
}

/**
 * Creates an action button for a specific workflow
 */
function createWorkflowButton(wf) {
  const btn = document.createElement("button");
  btn.className = "workflow-btn";
  btn.textContent = wf.name;
  btn.onclick = () => executeWorkflow(wf);
  return btn;
}

/**
 * Converts multiline template text into formatted HTML paragraphs,
 * preserving blank lines and paragraph breaks as they appear in the spreadsheet.
 */
function formatTemplateToHtml(plainText) {
  let text = plainText.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  text = text.replace(/\n[\u00A0\s]+\n/g, "\n\n");

  const paragraphs = text.split(/\n\s*\n/);

  return paragraphs
    .map((p) => {
      const trimmed = p.trim();
      if (!trimmed) return "";
      const withBreaks = trimmed.replace(/\n/g, "<br>");
      return `<p style="margin: 0 0 12pt 0; font-family: Calibri, Helvetica, Arial, sans-serif; font-size: 11pt; color: #000000; line-height: 1.4;">${withBreaks}</p>`;
    })
    .filter(Boolean)
    .join("");
}

/**
 * Converts Google Drive, OneDrive, or SharePoint URLs into direct file download streams
 */
function getDirectDownloadUrl(url) {
  if (!url) return "";
  let cleanUrl = url.trim();

  // If a custom filename was provided via pipe ("URL | Filename.pdf"), isolate URL
  if (cleanUrl.includes("|")) {
    cleanUrl = cleanUrl.split("|")[0].trim();
  }

  // 1. Google Drive Sharing Links (handles /file/d/ID/view, /open?id=ID, /uc?id=ID)
  if (cleanUrl.includes("drive.google.com") || cleanUrl.includes("docs.google.com")) {
    let fileId = "";
    const fileDMatch = cleanUrl.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
    const idParamMatch = cleanUrl.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    const dMatch = cleanUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);

    if (fileDMatch && fileDMatch[1]) {
      fileId = fileDMatch[1];
    } else if (idParamMatch && idParamMatch[1]) {
      fileId = idParamMatch[1];
    } else if (dMatch && dMatch[1]) {
      fileId = dMatch[1];
    }

    if (fileId) {
      return `https://drive.google.com/uc?export=download&id=${fileId}`;
    }
  }

  // 2. OneDrive / SharePoint Sharing Links
  if (cleanUrl.includes("sharepoint.com") || cleanUrl.includes("1drv.ms") || cleanUrl.includes("onedrive.live.com")) {
    if (!cleanUrl.includes("download=1")) {
      const separator = cleanUrl.includes("?") ? "&" : "?";
      return cleanUrl + separator + "download=1";
    }
  }

  return cleanUrl;
}

/**
 * Resolves a readable filename for the attachment (e.g. "Private Dining Menu.pdf")
 */
function getAttachmentFileName(pathOrUrl, workflowName) {
  if (!pathOrUrl) return "Private Dining Menu.pdf";

  // Optional manual override in sheet: "https://... | CustomName.pdf"
  if (pathOrUrl.includes("|")) {
    const parts = pathOrUrl.split("|");
    const customName = parts[1].trim();
    if (customName) {
      return customName.toLowerCase().endsWith(".pdf") ? customName : customName + ".pdf";
    }
  }

  // Direct file URL with filename
  try {
    const clean = pathOrUrl.split("?")[0];
    const parts = clean.split("/");
    const lastPart = parts[parts.length - 1];
    if (lastPart && lastPart.toLowerCase().endsWith(".pdf")) {
      return decodeURIComponent(lastPart);
    }
  } catch (e) {
    console.error("Filename parsing error:", e);
  }

  // Fallback based on workflow title
  if (workflowName) {
    const cleanName = workflowName.replace(/[^a-zA-Z0-9_\- ]/g, "").trim();
    if (cleanName.toLowerCase().includes("inquiry") || cleanName.toLowerCase().includes("response")) {
      return "Private Dining Menu.pdf";
    }
    return `${cleanName}.pdf`;
  }

  return "Private Dining Menu.pdf";
}

/**
 * Handles zero-click file attachment via Office.js
 */
function handleAttachment(workflow) {
  if (!workflow.attachmentPath) return;

  const rawPath = workflow.attachmentPath.trim();

  // If provided as a web URL (Google Drive, OneDrive, or direct link)
  if (rawPath.startsWith("http://") || rawPath.startsWith("https://")) {
    const downloadUrl = getDirectDownloadUrl(rawPath);
    const fileName = getAttachmentFileName(rawPath, workflow.name);

    showStatus(`Attaching ${fileName}...`, "info");

    Office.context.mailbox.item.addFileAttachmentAsync(
      downloadUrl,
      fileName,
      { isInline: false },
      function (attachResult) {
        if (attachResult.status === Office.AsyncResultStatus.Succeeded) {
          showStatus(`Injected text & attached "${fileName}".`, "success");
        } else {
          console.warn("Attachment error:", attachResult.error);
          showStatus(`Text injected. Attachment notice: ${attachResult.error.message}`, "warning");
        }
      }
    );
  } else {
    // If it's still a relative path like "/Menus/PrivateDining.pdf"
    showStatus(`Text injected. (To auto-attach, paste your Google Drive link into Column D of your spreadsheet).`, "warning");
  }
}

/**
 * Executes the workflow: parses email body, populates variables, injects HTML into draft, and attaches file
 */
function executeWorkflow(workflow) {
  Office.context.mailbox.item.body.getAsync(
    "html",
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        const rawEmailHtml = asyncResult.value;
        
        // 1. Parse client details from email
        const parsedVariables = parseInquiryEmail(rawEmailHtml);
        
        // 2. Build personalized text
        const finalText = buildFinalDraftText(workflow.template, parsedVariables);
        
        // 3. Convert text to HTML paragraphs to strictly preserve spacing and line breaks
        const finalHtml = formatTemplateToHtml(finalText);

        // 4. Prepend formatted HTML into the Outlook draft
        Office.context.mailbox.item.body.prependAsync(
          finalHtml,
          { coercionType: Office.CoercionType.Html },
          function (htmlResult) {
            if (htmlResult.status === Office.AsyncResultStatus.Succeeded) {
              showStatus(`Injected: ${workflow.name}`, "success");
              // 5. Trigger zero-click attachment
              handleAttachment(workflow);
            } else {
              // Fallback to text insertion if the email is plain text
              Office.context.mailbox.item.body.prependAsync(
                finalText + "\n\n",
                { coercionType: Office.CoercionType.Text },
                function (textResult) {
                  if (textResult.status === Office.AsyncResultStatus.Succeeded) {
                    showStatus(`Injected: ${workflow.name}`, "success");
                    handleAttachment(workflow);
                  } else {
                    console.error("Insertion failed:", textResult.error.message);
                    showStatus("Failed to insert text into draft.", "error");
                  }
                }
              );
            }
          }
        );
      } else {
        console.error("Failed to read email body.");
        showStatus("Could not read inquiry email.", "error");
      }
    }
  );
}

/**
 * Parses Toast inquiry emails based on verified HTML structure
 */
function parseInquiryEmail(emailHtml) {
  if (!emailHtml || typeof emailHtml !== "string") {
    return {
      firstName: "{{first_name}}",
      fullDate: "{{full_date}}",
      dayOfWeek: "{{day_of_week}}"
    };
  }

  // 1. Remove style and script blocks
  let cleanText = emailHtml
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, "")
    .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, "");

  // 2. Remove Outlook reply/forward header blocks and signatures
  cleanText = cleanText
    .replace(/<div id="divRplyFwdMsg"[\s\S]*?<\/div>/gi, "\n")
    .replace(/<div id="Signature"[\s\S]*?<\/div>/gi, "\n");

  // 3. Decode entities and convert structural HTML tags to standard newlines
  cleanText = cleanText
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/&nbsp;/gi, " ")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/(p|div|tr|td|li|h[1-6])>/gi, "\n")
    .replace(/<[^>]+>/g, " ")
    .replace(/[\u00A0\t]/g, " ")
    .replace(/[ ]{2,}/g, " ");

  // 4. Split into distinct trimmed lines
  const lines = cleanText
    .split("\n")
    .map((line) => line.trim())
    .filter((line) => line.length > 0);

  let firstName = "";
  let fullDate = "";
  let dayOfWeek = "";

  // Terms to prevent matching business names or locations
  const nameExclude = ["business", "organization", "company", "restaurant", "venue", "location"];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    // Skip standard metadata headers
    if (/^(from|to|sent|subject|cc|bcc):/i.test(line)) {
      continue;
    }

    // --- 1. Client First Name Extraction ---
    if (!firstName) {
      const isExcluded = nameExclude.some((term) => line.toLowerCase().includes(term));
      if (!isExcluded) {
        // Label on its own line: "Name" or "Name:"
        if (/^(?:guest\s+|customer\s+|client\s+)?name\s*[:]?$/i.test(line)) {
          if (i + 1 < lines.length && !lines[i + 1].includes(":")) {
            firstName = extractFirstName(lines[i + 1].trim());
          }
        }
        // Label inline: "Name: Robin Sorensen"
        else {
          const match = line.match(/^(?:guest\s+|customer\s+|client\s+)?name\s*[:\-]\s*(.+)$/i);
          if (match && match[1]) {
            firstName = extractFirstName(match[1].trim());
          }
        }
      }
    }

    // --- 2. Event Date Extraction ---
    if (!fullDate) {
      // Label on its own line: "Event date" or "Event date:"
      if (/^(?:event\s+date|date\s+of\s+event|requested\s+date|date)\s*[:]?$/i.test(line)) {
        if (i + 1 < lines.length && !lines[i + 1].includes(":")) {
          fullDate = cleanDate(lines[i + 1].trim());
        }
      }
      // Label inline: "Event date: Sunday, December 20, 2026"
      else {
        const match = line.match(/^(?:event\s+date|date\s+of\s+event|requested\s+date|date)\s*[:\-]\s*(.+)$/i);
        if (match && match[1]) {
          fullDate = cleanDate(match[1].trim());
        }
      }
    }
  }

  // --- 3. Day of Week Isolation ---
  if (fullDate) {
    const days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
    for (const d of days) {
      if (new RegExp("\\b" + d + "\\b", "i").test(fullDate)) {
        dayOfWeek = d;
        break;
      }
    }
    if (!dayOfWeek && fullDate.includes(",")) {
      dayOfWeek = fullDate.split(",")[0].trim();
    }
  }

  return {
    firstName: firstName || "{{first_name}}",
    fullDate: fullDate || "{{full_date}}",
    dayOfWeek: dayOfWeek || "{{day_of_week}}"
  };
}

function extractFirstName(nameStr) {
  const cleaned = nameStr.replace(/^(mr\.|ms\.|mrs\.|dr\.)\s+/i, "");
  return cleaned.split(/\s+/)[0].replace(/[^a-zA-Z'\-]/g, "");
}

function cleanDate(dateStr) {
  return dateStr.split("<")[0].replace(/\s*(at|@)\s*\d+.*$/i, "").trim();
}

/**
 * Replaces placeholders with parsed variables, preventing duplicate day names
 */
function buildFinalDraftText(templateText, variables) {
  let dateWithoutDay = variables.fullDate;
  if (variables.dayOfWeek && dateWithoutDay) {
    const dayPrefixRegex = new RegExp("^" + variables.dayOfWeek + "\\s*,?\\s*", "i");
    dateWithoutDay = dateWithoutDay.replace(dayPrefixRegex, "").trim();
  }

  let finalizedText = templateText
    .replace(/\{\{day_of_week\}\},\s*\{\{full_date\}\}/gi, variables.dayOfWeek + ", " + dateWithoutDay)
    .replace(/\{\{day_of_week\}\}\s+\{\{full_date\}\}/gi, variables.dayOfWeek + " " + dateWithoutDay)
    .replace(/\{\{first_name\}\}/g, variables.firstName)
    .replace(/\{\{full_date\}\}/g, variables.fullDate)
    .replace(/\{\{day_of_week\}\}/g, variables.dayOfWeek);

  return finalizedText;
}
