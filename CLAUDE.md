# Fiji Quote Calculator

## Project Structure
- src/Code.gs - Google Apps Script backend
- src/Index.html - Main HTML template
- src/JavaScript.html - All client-side JavaScript
- src/Stylesheet.html - All CSS styles

## Environments
- DEV Script ID: 1b_LE1KiiUZUtrOiNlP308BlO7y8pNNbBWZL0-7IOn7xuCV7DlmSKvBbn
- PROD Script ID: 1gWaXoz_gogPZ6wjr601NRyCy3UFgJwueN8dg04CDbj7Ce-J9miuYShMn
- DEV Spreadsheet: 1TAfQy6KRWLbBNZnU2TrnZuZ3BFd26HC0_sMqCePZRlo
- PROD Spreadsheet: 1YkNVq6StGbG_ZtlNICimWLUoRguePSQ_aOZrb_X-qSY

## Deploy Commands
DEV:
1. cp .clasp-dev.json .clasp.json && clasp push --force
2. Apps Script editor > Deploy > Manage deployments > Edit (pencil) > Version: New version > Save
   (URL stays the same — never use `clasp deploy` as it breaks the Web App deployment type)

PROD:
1. cp .clasp-prod.json .clasp.json && clasp push --force && cp .clasp-dev.json .clasp.json
2. Apps Script editor (PROD project) > Deploy > Manage deployments > Edit (pencil) > Version: New version > Save

GitHub: git add . && git commit -m "description" && git push

## DEV URL
https://script.google.com/a/macros/travelwithviva.com/s/AKfycbzXsrEfi_SpNmpGgElEXEGjbazjwfflLkc_NDthJco7YVHac2OZ2g3jo9SC6-4RdxOU2Q/exec

## PROD URL
https://script.google.com/a/macros/travelwithviva.com/s/AKfycbwaoQn3bYR_4nRb5sXTdbZTSMTSXf3n8UhCXtW9lKNwANmhFstiJLNom6iyJDdldBsE/exec