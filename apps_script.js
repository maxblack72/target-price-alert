// LONG TERM PRICE CHECK

const columnName                  = 0;
const columnTicker                = 1;
const columnTargetPrice           = 2;
const columnTargetPercentage      = 3;
const columnSignal                = 4; // Take Profit, Buy Limit, Stop Loss
const columnCurrentPrice          = 5;
const columnCurrentDiffPercentage = 6;
const columnSparkeline            = 7; // not used
const columnEmailSent             = 8;
const columnYFlink                = 9; // Link
const sheeetTab                   = "Prices";

let bodyMessage = '';
let emailSignal = '';
// a message may have multiple items/stock 
let sendMessage = false; 
let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheeetTab);

function addStrategyRowToEmail(row, i) {
  sendMessage = true;
  emailSignal = row[columnSignal].toLowerCase() || '';
  bodyMessage += `<strong>${row[columnName]}</strong>: there is a ${row[columnSignal].toUpperCase()} signal since 
    current price is €${row[columnCurrentPrice].toFixed(2)}<br>`;
  i++;
  // stop sending message after the first one
  sheet.getRange(`I${i}`).setValue(true);
}

function addPercentageRowToEmail(row, i) {
  sendMessage = true;
  emailSignal = row[columnSignal].toLowerCase() || '';
  bodyMessage += `<strong>${row[columnName]}</strong>: current price is €${row[columnCurrentPrice].toFixed(2)} 
    and is ${row[columnTargetPercentage]*100}% around your target price<br>`;
  i++;
  // stop sending message after the first one
  sheet.getRange(`I${i}`).setValue(true);
}

function checkPrices() {
  // Pulls data from the spreadsheet
  let source = sheet.getRange("A:I");
  let data = source.getValues();

  //Loops through the cells in the spreadsheet to find cells where the stock fell below purchase price
  let n = 0;
  for (let i in data) {
    //Skips the first row
    if (n++ == 0) continue;

    //Loads the current row
    let row = data[i];

    //Once at the end of the list, exits the loop
    if (row[columnTicker] == "") break;

    // if mail has been sent, go to next line
    if (row[columnEmailSent]) continue;

    // DEPENDING ON THE STRATEGY
    if (row[columnSignal] == 'Buy Limit' || row[columnSignal] == 'Stop Loss') {
      // if price is lower than target price
      if (row[columnCurrentPrice] <= row[columnTargetPrice]) {
        this.addStrategyRowToEmail(row, i);
        continue;
      }   
    } else if (row[columnSignal] == 'Take Profit') {
      // if price is upper than target price
      if (row[columnCurrentPrice] >= row[columnTargetPrice]) {
        this.addStrategyRowToEmail(row, i);
        continue;
      }
    }

    // if current percentage diff is around target percentage
    if (Math.abs(row[columnCurrentDiffPercentage]) < row[columnTargetPercentage]) {
        this.addPercentageRowToEmail(row, i);
        continue;
    }
  }

  if (!sendMessage) return;

  this.sendEmail(bodyMessage);
}

function sendEmail(bodyMessage) {
  const intro = "<h3>Price alert</h3>";
  // console.log(intro + bodyMessage)
  MailApp.sendEmail({
    to: SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail(),
    subject: "Target price alert: "+  emailSignal,
    htmlBody: intro + bodyMessage,
  });
}

// uncomment this line for local testing
module.exports = {
  checkPrices
}
