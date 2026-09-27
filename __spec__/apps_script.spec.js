// Import the function you want to test

const test1 = ['Etherium', 'ETH-EUR',  800, 4.00, 'Buy Limit', 1000, 1.8, false];
const test2 = ['Etherium', 'ETH-EUR',  1000, 4.00, 'Take Profit', 1500, 50, false];
const test3 = ['Etherium', 'ETH-EUR',  500, 4.00, 'Buy Limit', 1000, 50, false];
const test4 = ['Etherium', 'ETH-EUR',  500, 4.00, 'Buy Limit', 1000, 50, false];

global.SpreadsheetApp = {
  getActiveSpreadsheet: () => ({
    getOwner: () => ({
      getEmail: jest.fn((name) => null)
    }),
    getSheetByName: (name) => ({
      getRange: (range) => ({
        getValues: jest.fn((name) => [
          ['headings'],
          test1,
          test2,
          test3,
          test4
        ]),
        setValue: jest.fn((boolValue) => true)
      })
    }),
  }),
};

global.columnName                  = 0;
global.columnTicker                = 1;
global.columnTargetPrice           = 2;
global.columnTargetPercentage      = 3;
global.columnStrategy              = 4;
global.columnCurrentPrice          = 5;
global.columnCurrentDiffPercentage = 6;
global.columnEmailSent             = 7;

let {
  checkPrices
} = require('../apps_script');

// console.log(addPercentageRowToEmail.toString());

describe('checkPrices', () => {
  afterEach(() => {
    jest.restoreAllMocks();
  });
  beforeEach(() => {
    global.MailApp = {
      sendEmail: jest.fn((body) => null)
    };
	});

  test('should pull data from the "Prices" sheet', () => {
    // Call the checkPrices to be tested
    // console.log(addPercentageRowToEmail.toString());
    const mockAddPercentageRowToEmail = jest.fn((row, i) => null);
    addPercentageRowToEmail = mockAddPercentageRowToEmail;
    const mockAddStrategyRowToEmail = jest.fn((row, i) => null);
    addStrategyRowToEmail = mockAddStrategyRowToEmail;

    // addPercentageRowToEmail(); //esegue la moccata ( il test diventa ok)
    checkPrices(); // non esegue la moccata ma l'originale
    expect(mockAddPercentageRowToEmail).toHaveBeenCalledWith(test1, '1');
    expect(mockAddStrategyRowToEmail).toHaveBeenCalledWith(test2, '2');
  });

});

