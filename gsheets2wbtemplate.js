// Includes functions for exporting active sheet or all sheets as JSON object (also Python object syntax compatible).
// Tweak the makePrettyJSON_ function to customize what kind of JSON to export.

var FORMAT_ONELINE = 'One-line';
var FORMAT_MULTILINE = 'Multi-line';
var FORMAT_PRETTY = 'Pretty';

var LANGUAGE_JS = 'JavaScript';
var LANGUAGE_PYTHON = 'Python';

var STRUCTURE_LIST = 'List';
var STRUCTURE_HASH = 'Hash (keyed by "id" column)';

/* Defaults for this particular spreadsheet, change as desired */
var DEFAULT_FORMAT = FORMAT_PRETTY;
var DEFAULT_LANGUAGE = LANGUAGE_JS;
var DEFAULT_STRUCTURE = STRUCTURE_LIST;

function onOpen() {
  //creating new menu
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [
    //  {name: "Export JSON for this sheet", functionName: "exportSheet"},
    { name: 'Export JSON for all sheets', functionName: 'exportAllSheets' }
  ];
  ss.addMenu('Export JSON', menuEntries);
}

// function makeLabel(app, text, id) {
//   var lb = app.createLabel(text);
//   if (id) lb.setId(id);
//   return lb;
// }

// function makeListBox(app, name, items) {
//   var listBox = app.createListBox().setId(name).setName(name);
//   listBox.setVisibleItemCount(1);

//   var cache = CacheService.getPublicCache();
//   var selectedValue = cache.get(name);
//   Logger.log(selectedValue);
//   for (var i = 0; i < items.length; i++) {
//     listBox.addItem(items[i]);
//     if (items[1] == selectedValue) {
//       listBox.setSelectedIndex(i);
//     }
//   }
//   return listBox;
// }

// function makeButton(app, parent, name, callback) {
//   var button = app.createButton(name);
//   app.add(button);
//   var handler = app.createServerClickHandler(callback).addCallbackElement(parent);;
//   button.addClickHandler(handler);
//   return button;
// }

//For JSON output??
function makeTextBox(app, name) {
  var textArea = app.createTextArea().setWidth('100%').setHeight('200px').setId(name).setName(name);
  return textArea;
}

function exportAllSheets(e) {
  console.log('exportAllSheets');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets(); //Sheet[] — An array of all the sheets in the spreadsheet.

  var sheetsData = {};
  var convertOptions = getExportOptions(e);

  for (var i = 1; i < sheets.length; i++) {
    var sheet = sheets[i];
    var sheetName = sheet.getName();
    if (sheetName == 'parameters') {
      convertOptions.structure = STRUCTURE_HASH;
      //correctParId (sheet);
      //getColumnsData_(sheet, sheet.getRange(sheet.getFrozenRows()+1, 2, sheet.getMaxRows(), sheet.getMaxColumns()), 1);
    } else convertOptions.structure = STRUCTURE_LIST;

    var rowsData;
    if (sheetName == 'ru') {
      rowsData = getColumnsData_(
        sheet,
        sheet.getRange(sheet.getFrozenRows() + 1, 2, sheet.getMaxRows(), sheet.getMaxColumns()),
        1
      );
    } else {
      rowsData = getRowsData_(sheet, convertOptions);
    }

    //если нет заполненных строк кроме заголовка, то секцию не добаляем в шаблон
    console.log(Object.keys(rowsData).length);
    if (Object.keys(rowsData).length == 0) continue;

    sheetsData[sheetName] = rowsData;
    // console.log('Sheet name = ' + sheetName);
  }

  convertOptions.structure = STRUCTURE_LIST;
  var sheetMainData = getRowsData_(sheets[0], convertOptions);

  //sheetMainData = "123" + sheetMainData;
  //console.log(sheetMainData[0]);
  //console.log(typeof (sheetMainData[0]));

  //  sheetMainData.push(sheetsData);

  //  sheetsData["groups"] = sheetsData["groups"] + sheetMainData[0];

  // console.log(sheetsData["groups"]);
  // console.log(sheetsData["parameters"]);
  // //delete json[0];

  console.log(sheetsData);
  var templateBody = makeJSON_(sheetsData, getExportOptions(e));
  var templateHead = makeJSON_(sheetMainData, getExportOptions(e));

  //console.log(templateRawHead);
  templateHead = templateHead.replace('[', '');
  templateHead = templateHead.replace('"###",', '{');
  templateHead = templateHead.replace('    }', '');
  templateHead = templateHead.replace('\n\n]', '');
  templateHead += ',';
  //  templateHead += "\n";
  //  console.log(templateHead);

  // var templateHead2 = templateHead.replace("}", "");
  // //templateHead2 += ",";
  //   console.log(templateHead2);
  // templateHead += "123";
  // //templateHead[0] = "[";

  // var replacement = '[';
  // const testString = '12234345556';
  // console.log(testString);
  // newString = testString.replace('2', '!!!');
  // console.log(newString);
  // const p = 'The quick brown fox jumps over the lazy dog. If the dog reacted, was it really lazy?';

  // console.log(p.replace('dog', 'monkey'));

  //  var json = makeJSON_(sheetMainData, getExportOptions(e));
  //console.log('JSON elements');
  //delete json[0];
  //console.log(json[0]);

  // let str = "I love JavaScript";
  // let result = str.replace("I", "Oi");
  //console.log(result);

  //combain hader and bodyoftemplate
  var deviceTemplate = templateHead + templateBody;
  deviceTemplate = deviceTemplate.replace(',{', ',\n\n');
  //adding last }
  deviceTemplate += '\n}';

  //"[ replacewith [; ]"replace with ]
  deviceTemplate = deviceTemplate.replace(/\"\[/g, '[');
  deviceTemplate = deviceTemplate.replace(/\]\"/g, ']');

  //In parameters enum_titles replace \" to "
  deviceTemplate = deviceTemplate.replace(/\\\"/g, '"');

  //"true" -> true
  deviceTemplate = deviceTemplate.replace(/\"true\"/g, 'true');
  //"false" -> false
  deviceTemplate = deviceTemplate.replace(/\"false\"/g, 'false');

  var outputWindowHeader;
  try {
    var deviceTemplateParsed = JSON.parse(deviceTemplate);
    deviceTemplate = JSON.stringify(deviceTemplateParsed, null, 4);
    outputWindowHeader = 'JSON created successfully';
  } catch (err) {
    console.log(err.name); // ReferenceError
    console.log(err.message); // lalala is not defined
    console.log(err.stack); // ReferenceError: lalala is not defined at (...стек вызовов)

    // Можем также просто вывести ошибку целиком
    // Ошибка приводится к строке вида "name: message"
    console.log(err);
    outputWindowHeader = 'Exported with ERRORS!!! Check JSON data!!!';
  }

  //   deviceTemplate = JSON.stringify(JSON.parse(deviceTemplate),null,4);

  //console.log(deviceTemplate);
  //displayText_(deviceTemplate);
  displayText_(deviceTemplate, outputWindowHeader);
}

function exportSheet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  var rowsData = getRowsData_(sheet, getExportOptions(e));
  var json = makeJSON_(rowsData, getExportOptions(e));

  displayText_(json, 'Exported without Errors');
}

function correctParId(sheetParameters) {
  var parIdRange = sheetParameters.getRange(
    sheetParameters.getFrozenRows() + 1,
    1,
    sheetParameters.getMaxRows(),
    sheetParameters.getMaxColumns()
  );

  console.log(sheetParameters.getFrozenRows() + 1);
  console.log(sheetParameters.getMaxRows());
  console.log(parIdRange[0]);

  console.log(getColumnsData_(sheetParameters, parIdRange, 1));
}

function getExportOptions(e) {
  var options = {};

  options.language = (e && e.parameter.language) || DEFAULT_LANGUAGE;
  options.format = (e && e.parameter.format) || DEFAULT_FORMAT;
  options.structure = (e && e.parameter.structure) || DEFAULT_STRUCTURE;

  var cache = CacheService.getPublicCache();
  cache.put('language', options.language);
  cache.put('format', options.format);
  cache.put('structure', options.structure);

  Logger.log(options);
  return options;
}

// function replacer(key, value) {
//   if (typeof value == "\[") {
//     return value.toString()
//   }
//   return value
// }

function makeJSON_(object, options) {
  if (options.format == FORMAT_PRETTY) {
    var jsonString = JSON.stringify(object, null, 4);
  } else if (options.format == FORMAT_MULTILINE) {
    var jsonString = Utilities.jsonStringify(object);
    jsonString = jsonString.replace(/},/gi, '},\n');
    jsonString = jsonString.replace(/":\[{"/gi, '":\n[{"');
    jsonString = jsonString.replace(/}\],/gi, '}],\n');
  } else {
    var jsonString = Utilities.jsonStringify(object);
  }
  if (options.language == LANGUAGE_PYTHON) {
    // add unicode markers
    jsonString = jsonString.replace(/"([a-zA-Z]*)":\s+"/gi, '"$1": u"');
  }
  console.log('JSON string', jsonString);
  return jsonString;
}

//View result in JSON
function displayText_(text, windowHeader) {
  var output = HtmlService.createHtmlOutput(
    "<textarea style='width:100%;' rows='20'>" + text + '</textarea>'
  );
  output.setWidth(600);
  output.setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(output, windowHeader);
}

// getRowsData iterates row by row in the input range and returns an array of objects.
// Each object contains all the data for a given row, indexed by its normalized column name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//   - columnHeadersRowIndex: specifies the row number where the column names are stored.
//       This argument is optional and it defaults to the row immediately above range;
// Returns an Array of objects.
function getRowsData_(sheet, options) {
  var headersRange = sheet.getRange(1, 1, sheet.getFrozenRows(), sheet.getMaxColumns());
  var headers = headersRange.getValues()[0];
  var dataRange = sheet.getRange(
    sheet.getFrozenRows() + 1,
    1,
    sheet.getMaxRows(),
    sheet.getMaxColumns()
  );

  var objects = getObjects_(dataRange.getValues(), normalizeHeaders_(headers));
  if (options.structure == STRUCTURE_HASH) {
    var objectsById = {};
    objects.forEach(function (object) {
      objectsById[object.id] = object;
      delete objectsById[object.id].id;
    });

    return objectsById;
  } else {
    return objects;
  }
}

// getColumnsData iterates column by column in the input range and returns an array of objects.
// Each object contains all the data for a given column, indexed by its normalized row name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//   - rowHeadersColumnIndex: specifies the column number where the row names are stored.
//       This argument is optional and it defaults to the column immediately left of the range;
// Returns an Array of objects.
function getColumnsData_(sheet, range, rowHeadersColumnIndex) {
  rowHeadersColumnIndex = rowHeadersColumnIndex || range.getColumnIndex() - 1;
  var headersTmp = sheet
    .getRange(range.getRow(), rowHeadersColumnIndex, range.getNumRows(), 1)
    .getValues();
  var headers = headersTmp; //normalizeHeaders_(arrayTranspose_(headersTmp)[0]);
  return getObjects_(arrayTranspose_(range.getValues()), headers);
}

// For every row of data in data, generates an object that contains the data. Names of
// object fields are defined in keys.
// Arguments:
//   - data: JavaScript 2d array
//   - keys: Array of Strings that define the property names for the objects to create
function getObjects_(data, keys) {
  var objects = [];
  for (var i = 0; i < data.length; ++i) {
    var object = {};
    var hasData = false;
    for (var j = 0; j < data[i].length; ++j) {
      var cellData = data[i][j];

      cellData = normalizeDataCell_(cellData, keys[j]);

      if (isCellEmpty_(cellData)) {
        continue;
      }
      object[keys[j]] = cellData;
      hasData = true;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}

// Returns an Array of normalized Strings.
// Arguments:
//   - headers: Array of Strings to normalize
function normalizeHeaders_(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    var key = normalizeHeaderCell_(headers[i]);
    if (key.length > 0) {
      keys.push(key);
    }
  }
  return keys;
}

// Normalizes a string, by removing all alphanumeric characters and using mixed case
// to separate words. The output will always start with a lower case letter.
// This function is designed to produce JavaScript object property names.
// Arguments:
//   - header: string to normalize
// Examples:
//   "First Name" -> "firstName"
//   "Market Cap (millions) -> "marketCapMillions
//   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
function normalizeHeaderCell_(header) {
  var key = ''; //key - key field of JSON
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == ' ' && key.length > 0) {
      upperCase = true;
      continue;
    }

    //Allowed symbols are letters, numbers, _
    if (!isAlnum_(letter) && letter != '_') {
      continue;
    }
    if (key.length == 0 && isDigit_(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  return key;
}

function normalizeDataCell_(cellRawData, key) {
  // var cellData = "";   //cellData - data field of JSON
  //убираем лишние пробелы из содержимого ячейки (перед данными, дублирующие, в конце строки)

  if (typeof cellRawData != 'number') {
    cellData = cellRawData.toString().trim();
  } else {
    cellData = cellRawData;
  }

  // var upperCase = false;
  // for (var i = 0; i < cellRawData.length; ++i) {
  //   var letter = cellRawData[i];
  //   var previousLetter = cellRawData[i - 1];
  //   if ((( (previousLetter == " ") && (cellData.length > 0)) || (i == 0)) && (key == "name")) {
  //     upperCase = true;
  //   //  continue;
  //   }

  //   // //Allowed symbols are letters, numbers, _
  //   // if ( (!isAlnum_(letter)) && (letter != "_") ) {
  //   //   continue;
  //   // }

  //   // if (key.length == 0 && isDigit_(letter)) {
  //   //   continue; // first character must be a letter
  //   // }

  //   if (upperCase) {
  //     upperCase = false;
  //     cellData += letter.toUpperCase();
  //   } else {
  //     cellData += letter.toLowerCase();
  //   }
  // }
  return cellData;
}

// Returns true if the cell where cellData was read from is empty.
// Arguments:
//   - cellData: string
function isCellEmpty_(cellData) {
  return typeof cellData == 'string' && cellData == '';
}

// Returns true if the character char is alphabetical or number, false otherwise.
function isAlnum_(char) {
  return (char >= 'A' && char <= 'Z') || (char >= 'a' && char <= 'z') || isDigit_(char);
}

// Returns true if the character char is a digit, false otherwise.
function isDigit_(char) {
  return char >= '0' && char <= '9';
}

// Given a JavaScript 2d Array, this function returns the transposed table.
// Arguments:
//   - data: JavaScript 2d Array
// Returns a JavaScript 2d Array
// Example: arrayTranspose([[1,2,3],[4,5,6]]) returns [[1,4],[2,5],[3,6]].
function arrayTranspose_(data) {
  if (data.length == 0 || data[0].length == 0) {
    return null;
  }

  var ret = [];
  for (var i = 0; i < data[0].length; ++i) {
    ret.push([]);
  }

  for (var i = 0; i < data.length; ++i) {
    for (var j = 0; j < data[i].length; ++j) {
      ret[j][i] = data[i][j];
    }
  }

  return ret;
}
