/**
 * Main function that runs when the web app URL is accessed.
 * OPTIMIZED: Splits data loading to prevent HTTP 500 Errors.
 */
 function getSpreadsheet() {
  const DEV_SCRIPT_ID = '1b_LE1KiiUZUtrOiNlP308BlO7y8pNNbBWZL0-7IOn7xuCV7DlmSKvBbn';
  const SPREADSHEET_ID = ScriptApp.getScriptId() === DEV_SCRIPT_ID
    ? '1TAfQy6KRWLbBNZnU2TrnZuZ3BFd26HC0_sMqCePZRlo'  // DEV sheet
    : '1YkNVq6StGbG_ZtlNICimWLUoRguePSQ_aOZrb_X-qSY'; // PROD sheet
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}
function doGet(e) {
  // JSON rates API for wholesale portal
  if (e.parameter.action === 'getRates') {
    return getRatesJson();
  }

  try {
    const ss = getSpreadsheet();

    const quoteId = e.parameter.q;
    let initialQuote = null;
    let relevantHotels = new Set();
    let isQuoteView = false;

    // 1. If viewing a specific quote, we load specific data
    if (quoteId) {
       isQuoteView = true;
       initialQuote = getQuote(quoteId);
       
       if (initialQuote && initialQuote.hotels) {
          initialQuote.hotels.forEach(h => {
             if(h.name && h.name !== 'Other') relevantHotels.add(h.name);
          });
       }
    }

    // Helper function to fetch sheet data
    const fetchAndFilter = (sheetName, filterKey) => {
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) return [];
        
        // CRITICAL FIX: If we are in "Agent Mode" (No Quote ID), do NOT load heavy data here.
        // We will load it asynchronously after the page loads to prevent timeouts.
        if (!isQuoteView) {
            return []; 
        }

        const data = sheetToObjects(sheet);
       
        // If Quote Mode, filter strictly
        if (relevantHotels.size > 0) {
            return data.filter(item => relevantHotels.has(item[filterKey]));
        }
        
        return [];
    };

    // Load Data - In Agent Mode, these return empty [] to keep the page load fast
    const ratesData = fetchAndFilter('RatesEntry', 'Hotel');
    const conditionsData = fetchAndFilter('Conditions', 'Hotel');
    const mealPlansData = fetchAndFilter('MealPlans', 'Hotel');
    const specialsData = fetchAndFilter('Specials', 'Hotel');
   
    // Lightweight sheets - Load these immediately
    const transfersSheet = ss.getSheetByName('Transfers');
    const transfersData = transfersSheet ? sheetToObjects(transfersSheet) : [];
   
    const consultantSheet = ss.getSheetByName('Consultants');
    const consultantsData = consultantSheet ? sheetToObjects(consultantSheet) : [];
   
    const masterCalcsSheet = ss.getSheetByName('MasterCalcs');
    const masterCalcsData = masterCalcsSheet ? getMasterCalcs(masterCalcsSheet) : {};
    masterCalcsData.TERMS_AND_CONDITIONS = 'https://www.vivatravel.au/about-us/terms-and-conditions/';
    masterCalcsData.VIVA_LOGO_URL = 'https://vivatravel.au/wp-content/uploads/Layer-2.svg';

    // Create the payload
    const payload = {
      isLiteMode: !isQuoteView, // Flag to tell front-end to fetch the rest of the data
      rates: ratesData,
      conditions: conditionsData,
      mealPlans: mealPlansData,
      specials: specialsData,
      consultants: consultantsData,
      masterCalcs: masterCalcsData,
      transfers: transfersData,
      url: ScriptApp.getService().getUrl(),
      requestParams: e.parameter,
      initialQuote: initialQuote
    };

    const htmlTemplate = HtmlService.createTemplateFromFile('Index');
    htmlTemplate.data = JSON.stringify(payload);

    return htmlTemplate.evaluate()
      .setTitle('Fiji Quote Calculator')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (error) {
    Logger.log(error.toString());
    return HtmlService.createHtmlOutput(`<p>An error occurred: ${error.toString()}</p>`);
  }
}

/**
 * NEW FUNCTION: Fetches the heavy data asynchronously.
 * This runs AFTER the page has loaded, preventing the 500 error.
 */
function getAsyncData() {
  const ss = getSpreadsheet();
  
  // Helper to fetch full sheet
  const fetchFull = (sheetName) => {
      const sheet = ss.getSheetByName(sheetName);
      return sheet ? sheetToObjects(sheet) : [];
  };

  return {
      rates: fetchFull('RatesEntry'),
      conditions: fetchFull('Conditions'),
      mealPlans: fetchFull('MealPlans'),
      specials: fetchFull('Specials')
  };
}

/**
 * Fetches key-value pairs from the MasterCalcs sheet.
 */
function getMasterCalcs(sheet) {
  const data = sheet.getDataRange().getValues();
  const calcs = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const key = row[0];
    const value = row[1];
    if (key) {
      calcs[key] = value;
    }
  }
  return calcs;
}

/**
 * Helper function to convert a Google Sheet's data into an array of objects.
 */
function sheetToObjects(sheet) {
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  return data.map(row => {
    const obj = {};
    headers.forEach((header, i) => {
      const cleanHeader = header.toString().trim();
      if (cleanHeader) {
        if (row[i] instanceof Date) {
          obj[cleanHeader] = row[i].toISOString();
        } else {
          obj[cleanHeader] = row[i];
        }
      }
    });
    return obj;
  });
}

/**
 * Saves a quote to the 'SavedQuotes' sheet.
 */
function saveQuote(quoteDetails) {
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName('SavedQuotes');
    if (!sheet) {
       // Create sheet if it doesn't exist to prevent crash
       sheet = ss.insertSheet('SavedQuotes');
       sheet.appendRow(['QuoteID', 'ClientName', 'DateSaved', 'QuoteData', 'Status', 'AcceptedDate', 'OptionSelected']);
    }

    const lastRow = sheet.getLastRow();
    let newIdNumber = 1;

    if (lastRow >= 2) {
      const lastId = sheet.getRange(lastRow, 1).getValue();
      const lastNumber = parseInt(String(lastId).replace(/\D/g, ''), 10);
      if (!isNaN(lastNumber)) {
        newIdNumber = lastNumber + 1;
      }
    }
    const quoteId = "Viva" + String(newIdNumber).padStart(4, '0');
    const clientName = quoteDetails.clientName || 'Unnamed Client';
    const dateSaved = new Date();
    const quoteDataString = JSON.stringify(quoteDetails.quoteData);

    // Initial save sets Status to "Pending" and leaves AcceptedDate/OptionSelected blank
    sheet.appendRow([quoteId, clientName, dateSaved, quoteDataString, "Pending", "", ""]);

    return quoteId;
  } catch (error) {
    Logger.log('Error in saveQuote: ' + error.toString());
    return { error: error.toString() };
  }
}

/**
 * Searches for saved quotes.
 */
function searchQuotes(searchTerm) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('SavedQuotes');

    if (!sheet) return [];

    const dataRange = sheet.getDataRange();
    if (dataRange.getNumRows() <= 1) return [];

    const data = dataRange.getValues();
    const headers = data.shift();

    const quoteIdIndex = headers.indexOf('QuoteID');
    const clientNameIndex = headers.indexOf('ClientName');
    const dateSavedIndex = headers.indexOf('DateSaved');

    if (quoteIdIndex === -1 || clientNameIndex === -1 || dateSavedIndex === -1) {
      // Graceful fail
      return [];
    }

    const allQuotes = [];
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (!row || !row[quoteIdIndex] || String(row[quoteIdIndex]).trim() === '') continue;
      try {
        allQuotes.push({
          quoteId: row[quoteIdIndex],
          clientName: row[clientNameIndex],
          dateSaved: new Date(row[dateSavedIndex]).toISOString()
        });
      } catch (e) {
        Logger.log(`Error processing row ${i+2}: ${row}. Error: ${e.message}`);
      }
    }

    if (!searchTerm || searchTerm.trim() === '') {
      return allQuotes.reverse();
    }

    const lowerCaseSearchTerm = searchTerm.toLowerCase();
    const filteredQuotes = allQuotes.filter(quote => {
      const quoteId = quote.quoteId ? String(quote.quoteId).toLowerCase() : '';
      const clientName = quote.clientName ? String(quote.clientName).toLowerCase() : '';
      return quoteId.includes(lowerCaseSearchTerm) || clientName.includes(lowerCaseSearchTerm);
    });

    return filteredQuotes.reverse();

  } catch (error) {
    Logger.log('Error in searchQuotes: ' + error.toString());
    return { error: 'An error occurred: ' + error.message };
  }
}

/**
 * Retrieves the full data for a specific quote by its ID.
 * FIX APPLIED: Converts Date objects to Strings to prevent transport errors on Accepted quotes.
 */
function getQuote(quoteId) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('SavedQuotes');
    if (!sheet) {
      return { error: "Sheet 'SavedQuotes' not found." };
    }
   
    const data = sheet.getDataRange().getValues();
    if (data.length === 0) return null;
   
    // Normalize headers: Trim spaces and handle potential empty headers
    const headers = data[0].map(h => h ? String(h).trim() : "");
   
    // Locate columns dynamically
    const quoteIdIndex = headers.indexOf('QuoteID');
    const quoteDataIndex = headers.indexOf('QuoteData');
    const statusIndex = headers.indexOf('Status');
    const acceptedDateIndex = headers.indexOf('AcceptedDate');
    const optionSelectedIndex = headers.indexOf('OptionSelected');

    // Basic validation
    if (quoteIdIndex === -1 || quoteDataIndex === -1) {
      return { error: `Required columns (QuoteID, QuoteData) are missing. Found headers: ${headers.join(', ')}` };
    }

    // Normalize the search ID (remove spaces, case insensitive)
    const targetId = String(quoteId).trim().toLowerCase();

    // Start loop from 1 (skipping header)
    for (let i = 1; i < data.length; i++) {
      // Robustly get the Row ID, handling numbers or text
      const cellValue = data[i][quoteIdIndex];
      const rowId = (cellValue === undefined || cellValue === null) ? "" : String(cellValue).trim().toLowerCase();
     
      if (rowId === targetId) {
        const rawJson = data[i][quoteDataIndex];

        // Check for empty data before parsing
        if (!rawJson || String(rawJson).trim() === "") {
           return { error: `Quote ${quoteId} found, but 'QuoteData' column is empty.` };
        }

        try {
          let quoteData = JSON.parse(rawJson);
         
          // --- Inject Metadata into the object ---
          // FIX: Default to "Pending" if status is missing/empty
          if (statusIndex > -1 && data[i].length > statusIndex) {
             quoteData.savedStatus = data[i][statusIndex] || "Pending";
          } else {
             quoteData.savedStatus = "Pending";
          }
         
          // FIX: Explicitly convert Date objects to Strings.
          // Apps Script can fail to serialize Date objects nested in objects returned via google.script.run
          if (acceptedDateIndex > -1 && data[i].length > acceptedDateIndex) {
             const dateVal = data[i][acceptedDateIndex];
             if (dateVal instanceof Date) {
                 quoteData.acceptedDate = dateVal.toISOString();
             } else {
                 quoteData.acceptedDate = dateVal;
             }
          }
         
          if (optionSelectedIndex > -1 && data[i].length > optionSelectedIndex) {
             quoteData.optionSelected = data[i][optionSelectedIndex];
          }

          return quoteData;
        } catch (jsonError) {
          return { error: "Data corrupted (JSON Parse Error): " + jsonError.message };
        }
      }
    }
    // If we exit the loop, ID wasn't found
    return { error: `Quote ID '${quoteId}' not found in ${data.length - 1} rows scanned.` };
  } catch (error) {
    Logger.log('Error in getQuote: ' + error.toString());
    return { error: error.toString() };
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * JSON API endpoint for the wholesale agent portal.
 * Called via: ?action=getRates
 * Returns all rate data needed to build quotes.
 */
function getRatesJson() {
  try {
    const ss = getSpreadsheet();
    const data = {
      rates:      sheetToObjects(ss.getSheetByName('RatesEntry'))  || [],
      conditions: sheetToObjects(ss.getSheetByName('Conditions'))  || [],
      mealPlans:  sheetToObjects(ss.getSheetByName('MealPlans'))   || [],
      specials:   sheetToObjects(ss.getSheetByName('Specials'))    || [],
      transfers:  sheetToObjects(ss.getSheetByName('Transfers'))   || [],
      masterCalcs: getMasterCalcs(ss.getSheetByName('MasterCalcs'))
    };
    return ContentService
      .createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Updates a quote status to 'Accepted' and sends a notification email.
 */
function acceptQuote(quoteId, option, price) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('SavedQuotes');
    if (!sheet) return { error: "Sheet not found" };
   
    const data = sheet.getDataRange().getValues();
    if (data.length === 0) return { error: "Sheet is empty" };

    // Normalize headers: Trim spaces
    const headers = data[0].map(h => h ? String(h).trim() : "");
   
    const quoteIdIndex = headers.indexOf('QuoteID');
    const quoteDataIndex = headers.indexOf('QuoteData');
    const statusIndex = headers.indexOf('Status');
    const acceptedDateIndex = headers.indexOf('AcceptedDate');
    const optionSelectedIndex = headers.indexOf('OptionSelected');
    const dateSavedIndex = headers.indexOf('DateSaved');
    const optionPriceIndex = headers.indexOf('OptionPrice'); // NEW: Find OptionPrice column
   
    if (quoteIdIndex === -1) return { error: "QuoteID column not found." };
    if (statusIndex === -1 || acceptedDateIndex === -1 || optionSelectedIndex === -1 || dateSavedIndex === -1) {
       return { error: "One or more required columns (Status, AcceptedDate, OptionSelected, DateSaved) are missing." };
    }
    // Check if OptionPrice column exists
    if (optionPriceIndex === -1) {
       return { error: "Column 'OptionPrice' not found in SavedQuotes sheet. Please add it." };
    }

    const optionToRecord = (option === 'Instant') ? 'Instant Purchase' : option;
    const targetId = String(quoteId).trim().toLowerCase();

    // Loop through rows (start at 1 to skip header)
    for (let i = 1; i < data.length; i++) {
      const cellValue = data[i][quoteIdIndex];
      const rowId = (cellValue === undefined || cellValue === null) ? "" : String(cellValue).trim().toLowerCase();

      if (rowId === targetId) {
        const acceptanceDate = new Date();
       
        // --- 48 HOUR CHECK LOGIC ---
        let isExpired = false;
        const dateSavedVal = data[i][dateSavedIndex];
        if (dateSavedVal) {
            const dateSaved = new Date(dateSavedVal);
            const timeDiff = acceptanceDate.getTime() - dateSaved.getTime();
            const hoursDiff = timeDiff / (1000 * 3600);
            if (hoursDiff > 48) {
                isExpired = true;
            }
        }

        // Update columns
        sheet.getRange(i + 1, statusIndex + 1).setValue("Accepted");
        sheet.getRange(i + 1, acceptedDateIndex + 1).setValue(acceptanceDate);
        sheet.getRange(i + 1, optionSelectedIndex + 1).setValue(optionToRecord);
        sheet.getRange(i + 1, optionPriceIndex + 1).setValue(price); // NEW: Save Price
       
        SpreadsheetApp.flush(); // Force update immediately

        // --- Email Notification Logic ---
        try {
          // 1. Extract Details from stored JSON
          let clientNames = "Valued Client";
          let displayPrice = price || "N/A"; // NEW: Use the passed price for the email
          let consultantName = "N/A";
         
          if (quoteDataIndex > -1) {
            const rawJson = data[i][quoteDataIndex];
            if (rawJson) {
              try {
                const quoteData = JSON.parse(rawJson);
                // Get Consultant Name
                if (quoteData.overview && quoteData.overview.consultant) {
                    consultantName = quoteData.overview.consultant;
                }
                // Get Client Names
                if (quoteData.overview && quoteData.overview.passengers && Array.isArray(quoteData.overview.passengers)) {
                   const names = quoteData.overview.passengers
                                   .map(p => p.name)
                                   .filter(n => n && n.trim() !== "");
                   if (names.length > 0) clientNames = names.join(", ");
                } else {
                   const clientNameIndex = headers.indexOf('ClientName');
                   if (clientNameIndex > -1) {
                      clientNames = data[i][clientNameIndex];
                   }
                }
                // Fallback price if passed price is empty (shouldn't happen with new code)
                if ((!price || price === "N/A") && quoteData.summaryPrices && quoteData.summaryPrices.total) {
                  displayPrice = quoteData.summaryPrices.total;
                }
              } catch (parseErr) {
                Logger.log("Error parsing JSON for email: " + parseErr);
              }
            }
          }

          // 2. Format Date
          const formattedDate = Utilities.formatDate(acceptanceDate, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy h:mm a");

          // 3. Construct Email
          const emailRecipient = "info@travelwithviva.com";
         
          let subject = `Quote Accepted: ${quoteId} - ${clientNames}`;
          if (isExpired) {
              subject = `[EXPIRED > 48 HRS] Quote Accepted: ${quoteId} - ${clientNames}`;
          }

          const expirationWarning = isExpired
            ? `<div style="background-color: #fff3cd; border: 1px solid #ffeeba; color: #856404; padding: 10px; margin-bottom: 15px; border-radius: 4px;">
                 <strong>⚠️ ATTENTION:</strong> This quote was created more than 48 hours ago. <br>
                 Prices and availability may have changed. Please reconfirm rates with suppliers immediately before processing.
               </div>`
            : '';
         
          const htmlBody = `
            <div style="font-family: Arial, sans-serif; color: #333;">
              <h2 style="color: ${isExpired ? '#d9534f' : '#2E86C1'};">
                  ${isExpired ? 'Quote Accepted (Expired)' : 'New Quote Acceptance'}
              </h2>
              ${expirationWarning}
              <p>A client has accepted a quote via the app.</p>
              <table style="border-collapse: collapse; width: 100%; max-width: 600px;">
                <tr>
                  <td style="padding: 8px; border-bottom: 1px solid #ddd; font-weight: bold;">Quote Reference:</td>
                  <td style="padding: 8px; border-bottom: 1px solid #ddd;">${quoteId}</td>
                </tr>
                <tr>
                  <td style="padding: 8px; border-bottom: 1px solid #ddd; font-weight: bold;">Consultant:</td>
                  <td style="padding: 8px; border-bottom: 1px solid #ddd;">${consultantName}</td>
                </tr>
                <tr>
                  <td style="padding: 8px; border-bottom: 1px solid #ddd; font-weight: bold;">Client Name(s):</td>
                  <td style="padding: 8px; border-bottom: 1px solid #ddd;">${clientNames}</td>
                </tr>
                <tr>
                  <td style="padding: 8px; border-bottom: 1px solid #ddd; font-weight: bold;">Option Selected:</td>
                  <td style="padding: 8px; border-bottom: 1px solid #ddd; color: ${optionToRecord === 'Instant Purchase' ? '#16a34a' : '#2563eb'}; font-weight: bold;">${optionToRecord}</td>
                </tr>
                 <tr>
                  <td style="padding: 8px; border-bottom: 1px solid #ddd; font-weight: bold;">Accepted Price:</td>
                  <td style="padding: 8px; border-bottom: 1px solid #ddd;">${displayPrice}</td>
                </tr>
                <tr>
                  <td style="padding: 8px; border-bottom: 1px solid #ddd; font-weight: bold;">Terms & Conditions:</td>
                  <td style="padding: 8px; border-bottom: 1px solid #ddd; color: #16a34a; font-weight: bold;">ACCEPTED</td>
                </tr>
                <tr>
                  <td style="padding: 8px; border-bottom: 1px solid #ddd; font-weight: bold;">Date Accepted:</td>
                  <td style="padding: 8px; border-bottom: 1px solid #ddd;">${formattedDate}</td>
                </tr>
              </table>
              <p style="margin-top: 20px; font-size: 12px; color: #666;">This is an automated message from the Fiji Quote Calculator App.</p>
            </div>
          `;

          GmailApp.sendEmail(emailRecipient, subject, "", {
            htmlBody: htmlBody,
            name: "Fiji Quote App"
          });

        } catch (emailError) {
          Logger.log("Failed to send acceptance email: " + emailError.toString());
        }

        return { success: true, expired: isExpired };
      }
    }
    return { error: "Quote ID not found." };
  } catch (e) {
    return { error: e.toString() };
  }
}

/**
 * !!! IMPORTANT !!!
 * RUN THIS FUNCTION ONCE MANUALLY TO AUTHORIZE EMAILS.
 * * 1. Select 'AUTHORIZE_EMAIL_HERE' from the dropdown menu at the top.
 * 2. Click 'Run'.
 * 3. Grant the permissions when asked.
 */
function AUTHORIZE_EMAIL_HERE() {
  const email = Session.getActiveUser().getEmail();
  GmailApp.sendEmail(email, "Test Email", "If you received this, the app is authorized to send emails.");
  Logger.log("Email authorization successful. You can now use the app.");
}

function testFetchAuth() {
  try {
    const response = UrlFetchApp.fetch('https://vivatravel.au/wp-content/uploads/Layer-2.svg', { muteHttpExceptions: true });
    Logger.log('SUCCESS - status: ' + response.getResponseCode() + ', type: ' + response.getBlob().getContentType());
  } catch(e) {
    Logger.log('FAILED: ' + e);
  }
}

function getImageAsBase64(url) {
  try {
    let blob;

    const driveId = (
      url.match(/drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)/) ||
      url.match(/drive\.google\.com\/open\?id=([a-zA-Z0-9_-]+)/) ||
      url.match(/drive\.google\.com\/[^?]*[?&]id=([a-zA-Z0-9_-]+)/) ||
      url.match(/lh3\.googleusercontent\.com\/d\/([a-zA-Z0-9_-]+)/)
    )?.[1];

    if (driveId) {
      blob = DriveApp.getFileById(driveId).getBlob();
    } else {
      const response = UrlFetchApp.fetch(url, {
        muteHttpExceptions: true,
        followRedirects: true,
        headers: {
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
          'Accept': 'image/webp,image/apng,image/*,*/*;q=0.8'
        }
      });
      const code = response.getResponseCode();
      if (code !== 200) {
        Logger.log('getImageAsBase64 HTTP ' + code + ' for: ' + url);
        return null;
      }
      blob = response.getBlob();
    }

    const mimeType = blob.getContentType() || 'image/jpeg';
    if (!mimeType.startsWith('image/')) {
      Logger.log('getImageAsBase64 non-image type "' + mimeType + '" for: ' + url);
      return null;
    }
    return 'data:' + mimeType + ';base64,' + Utilities.base64Encode(blob.getBytes());
  } catch (e) {
    Logger.log('getImageAsBase64 error for ' + url + ': ' + e);
    return null;
  }
}

/**
 * RUN THIS IN APPS SCRIPT EDITOR TO SEE WHAT IMAGE LINKS ARE STORED IN THE CONDITIONS SHEET.
 * Select this function from the dropdown and click Run, then check Execution Log.
 */
function DEBUG_listConditionsImageLinks() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Conditions');
  if (!sheet) { Logger.log('No Conditions sheet found!'); return; }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const hotelIdx = headers.indexOf('Hotel');
  const imgIdx = headers.indexOf('ImageLink');

  Logger.log('=== Conditions Sheet: Hotel Names & Image Links ===');
  Logger.log('Hotel column index: ' + hotelIdx + ', ImageLink column index: ' + imgIdx);

  if (hotelIdx === -1) { Logger.log('ERROR: No "Hotel" column found. Headers: ' + headers.join(', ')); return; }
  if (imgIdx === -1) { Logger.log('ERROR: No "ImageLink" column found. Headers: ' + headers.join(', ')); return; }

  for (let i = 1; i < data.length; i++) {
    const hotelName = data[i][hotelIdx];
    const imageLink = data[i][imgIdx];
    if (hotelName) {
      Logger.log(`Row ${i+1}: Hotel="${hotelName}" | ImageLink="${imageLink || '(empty)'}"`);
    }
  }
  Logger.log('=== End ===');
}

/**
 * RUN THIS IN APPS SCRIPT EDITOR TO DIAGNOSE IMAGE LOADING ISSUES.
 * Select this function from the dropdown and click Run, then check Execution Log.
 */
function DEBUG_testHotelImage() {
  const testUrl = 'https://drive.google.com/file/d/1f3-BchxWdZz8h7guG_k1g-7KaiBoXzuM/view?usp=drive_link';

  Logger.log('=== Hotel Image Debug Test ===');
  Logger.log('URL: ' + testUrl);

  // Step 1: Extract Drive ID
  const driveId = (
    testUrl.match(/drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)/) ||
    testUrl.match(/drive\.google\.com\/open\?id=([a-zA-Z0-9_-]+)/)
  )?.[1];
  Logger.log('Extracted Drive ID: ' + driveId);

  if (!driveId) {
    Logger.log('FAIL: Could not extract Drive ID from URL');
    return;
  }

  // Step 2: Try to get file from Drive
  try {
    const file = DriveApp.getFileById(driveId);
    Logger.log('File name: ' + file.getName());
    Logger.log('File MIME type: ' + file.getMimeType());
    Logger.log('File size: ' + file.getSize() + ' bytes');
    Logger.log('File sharing: ' + file.getSharingAccess());

    const blob = file.getBlob();
    Logger.log('Blob content type: ' + blob.getContentType());
    Logger.log('Blob size: ' + blob.getBytes().length + ' bytes');
    Logger.log('SUCCESS: Image loaded from Drive OK');
  } catch (e) {
    Logger.log('FAIL: DriveApp error - ' + e.toString());
    Logger.log('This usually means the file ID is wrong or the script cannot access this file.');
  }
}

























