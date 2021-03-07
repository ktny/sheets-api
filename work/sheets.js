var {google} = require('googleapis');
var {OAuth2Client} = require('google-auth-library');
var util = require('util');

var SheetsHelper = function(accessToken) {
  var auth = new OAuth2Client();
  auth.credentials = {
    access_token: accessToken
  };
  this.service = google.sheets({version: 'v4', auth: auth});
};

module.exports = SheetsHelper;

var COLUMNS = [
  { field: 'id', header: 'ID' },
  { field: 'customerName', header: 'Customer Name'},
  { field: 'productCode', header: 'Product Code' },
  { field: 'unitsOrdered', header: 'Units Ordered' },
  { field: 'unitPrice', header: 'Unit Price' },
  { field: 'status', header: 'Status'}
];

function buildHeaderRowRequest(sheetId) {
  var cells = COLUMNS.map(function(column) {
    return {
      userEnteredValue: {
        stringValue: column.header,
      },
      userEnteredFormat: {
        textFormat: {
          bold: true
        }
      }
    }
  });
  return {
    updateCells: {
      start: {
        sheetId: sheetId,
        rowIndex: 0,
        columnIndex: 0,
      },
      rows: [
        {
          values: cells,
        }
      ],
      fields: 'userEnteredValue,userEnteredFormat.textFormat.bold',
    }
  }
}

SheetsHelper.prototype.createSpreadsheet = function(title, callback) {
  var self = this;
  var request = {
    resource: {
      properties: {
        title: title
      },
      sheets: [
        {
          properties: {
            title: 'Data',
            gridProperties: {
              columnCount: 6,
              frozenRowCount: 1
            }
          }
        }
      ]
    }
  };
  // スプレッドシートを作成
  self.service.spreadsheets.create(request, function(err, response) {
    if (err) {
      return callback(err);
    }
    var spreadsheet = response.data;
    var dataSheetId = spreadsheet.sheets[0].properties.sheetId;
    var requests = [
      buildHeaderRowRequest(dataSheetId)
    ];

    var request = {
      spreadsheetId: spreadsheet.spreadsheetId,
      resource: {
        requests: requests,
      }
    };
    // 作成したスプレッドシートを更新
    self.service.spreadsheets.batchUpdate(request, function(err, response) {
      if (err) {
        return callback(err);
      }
      return callback(null, spreadsheet);
    });
  });
}
