body = {
  "requests": [
    {
      "repeatCell": {
        "range": {
          "sheetId": '0',
          "startRowIndex": 0,
          "endRowIndex": 1,
          "startColumnIndex": 0,
          "endColumnIndex": 10,
        },
        "cell": {
          "userEnteredFormat": {
            "backgroundColor": {
              "red": 0.203,
              "green": 0.196,
              "blue": 0.239
            },
            "horizontalAlignment" : "CENTER",
            "textFormat": {
              "foregroundColor": {
                "red": 1.0,
                "green": 1.0,
                "blue": 1.0
              },
              "fontSize": 10,
            }
          }
        },
        "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
      }
    },
    {
      "updateSheetProperties": {
        "properties": {
          "sheetId": '',
          "gridProperties": {
            "frozenRowCount": 1
          }
        },
        "fields": "gridProperties.frozenRowCount"
      }
    }
  ]
}
