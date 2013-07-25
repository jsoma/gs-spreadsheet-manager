var ManagedSpreadsheet = function(key, options) {
  if(!options)
    options = {}
  this.spreadsheet = SpreadsheetApp.openById(key);
  this.worksheets = {};
  this.cache = !!options.cache;
}

ManagedSpreadsheet.prototype.moveToFolder = function(name, parentFolderName) {
  var file = DocsList.getFileById(this.spreadsheet.getId());
  var parent_folder, folder;
  
  if(parentFolderName) {
    try {
      parent_folder = DocsList.getFolder(parentFolderName);
    } catch(e) {
      parent_folder = DocsList.createFolder(parentFolderName);
    }
  }
  try {
    if(parent_folder)
      folder = parent_folder.getFolder(name);
    else
      folder = DocsList.getFolder(name);
  } catch(e) {
    if(parent_folder)
      folder = parent_folder.createFolder(name);
    else
      folder = DocsList.createFolder(name);
  }

  file.addToFolder(folder);
  file.removeFromFolder(DocsList.getRootFolder());
}

ManagedSpreadsheet.prototype.processAppends = function(action) {
  for(var key in this.worksheets) {
    this.worksheets[key].processAppends();
  };
}

ManagedSpreadsheet.prototype.flush = function(action) {
  for(var key in this.worksheets) {
    this.worksheets[key].flush();
  }
}

ManagedSpreadsheet.prototype.atomic = function(action) {
  // Copy the spreadsheet into a new spreadsheet
  this.fork();
  // Run operations
  action.call(this);
  // Commit cached values to the spreadsheet
  this.flush();
  // Move the worksheets back into the initial spreadsheet
  this.merge();
}

ManagedSpreadsheet.prototype.fork = function() {
  this.original_spreadsheet = this.spreadsheet;
  this.spreadsheet = this.spreadsheet.copy("Temporary " + this.spreadsheet.getName() + ", " + new Date());
  this.worksheets = {};
}

ManagedSpreadsheet.prototype.sheetCount = function() {
  return this.original_spreadsheet.getSheets().length;
}

ManagedSpreadsheet.prototype.getId = function(name) {
  return this.spreadsheet.getId();
}

ManagedSpreadsheet.prototype.merge = function() {
  // Can't have a spreadsheet without sheets, so we'll add in a placeholder
  var placeholder = this.original_spreadsheet.insertSheet("Placeholder: " + new Date());

  var that = this;
  
  // First we'll remove every sheet except the placeholder sheet
  this.original_spreadsheet.getSheets().forEach( function(sheet) {
    if(sheet.getSheetId() === placeholder.getSheetId())
      return;

    sheet.activate();
    that.original_spreadsheet.deleteActiveSheet();
  });

  this.spreadsheet.getSheets().forEach( function(sheet) {
    var new_sheet = sheet.copyTo(that.original_spreadsheet);
    new_sheet.setName(sheet.getName());
  })
  
  // Remove the placeholder sheet
  placeholder.activate();
  this.original_spreadsheet.deleteActiveSheet();
  
  // Remove the copied document
  var file = DocsList.getFileById(this.spreadsheet.getId());
  file.setTrashed(true);
  
  // Go back to the original spreadsheet
  this.spreadsheet = this.original_spreadsheet;
  this.worksheets = {};
}

ManagedSpreadsheet.prototype.sheet = function(name) {
  if(!this.worksheets[name]) {
    var worksheet = this.spreadsheet.getSheets().filter( function(sheet) {
      return sheet.getName() === name;
    })[0];
    
    if(!worksheet) {
      return;
    }
    this.worksheets[name] = new ManagedWorksheet(this.spreadsheet, name, {cache: this.cache, worksheet: worksheet});
  }
  return this.worksheets[name];
}

var ManagedWorksheet = function(spreadsheet, name, options) {
  if(!options)
    options = {};
  this.spreadsheet = spreadsheet;
  // Why this way? I don't know. .getSheetByName won't work.
  this.worksheet = options.worksheet || this.spreadsheet.getSheets().filter( function(sheet) {
    return sheet.getName() === name;
  })[0];
  if(options.cache) {
    var colCount = this.worksheet.getLastColumn();
    var rowCount = this.worksheet.getLastRow();
    var range = this.worksheet.getRange(1, 1, rowCount, colCount);
    this._cache = range.getValues();
  }
}

ManagedWorksheet.prototype.getId = function() {
  return this.worksheet.getSheetId();
}
  
ManagedWorksheet.prototype.columnNames = function() {
  if(!this._columnNames)
    this._columnNames = this.getRow(1);
  return this._columnNames;
}

ManagedWorksheet.prototype.duplicateAs = function(newName) {
  this.worksheet.activate();
  var newSheet = this.spreadsheet.duplicateActiveSheet();
  newSheet.setName(newName);
  return new ManagedWorksheet(this.spreadsheet, newName, { worksheet: newSheet });
}

ManagedWorksheet.prototype.copyTo = function(managed_spreadsheet) {
  var copied = this.worksheet.copyTo(managed_spreadsheet.spreadsheet);
  var name = this.worksheet.getName();
  copied.setName(name);
  return new ManagedWorksheet(managed_spreadsheet.spreadsheet, name, { worksheet: copied });
}

// data should be in the form {key: value, key: value}
ManagedWorksheet.prototype.append = function(data, batch) {
  var toAppend;
  if(data instanceof Array) {
    toAppend = data;
  } else {
    toAppend = this.objectToRow(data);
  }
  if(batch) {
    if(!this.appendQueue)
      this.appendQueue = []
    this.appendQueue.push(toAppend);
  } else {
    if(this._cache) {
      this._cache.push(toAppend);
    } else {
      this.worksheet.appendRow(toAppend);
    }
    this._indices = null;
    this._lastRow = null;
  }
}

ManagedWorksheet.prototype.flush = function() {
  if(!this._cache)
    return;
  var range = this.worksheet.getRange( 1, 1, this._cache.length, this._cache[0].length )
  range.setValues(this._cache);
}

ManagedWorksheet.prototype.processAppends = function() {
  if(!this.appendQueue || this.appendQueue.length == 0)
    return;
  
  if(this._cache) {
    for(var i=0;i<this.appendQueue.length;i++) {
      this._cache.push(this.appendQueue[i]);
    }
    return;
  }
    
  var newRow = this.getLastRow() + 1;
  var range = this.worksheet.getRange(newRow, 1, this.appendQueue.length, this.appendQueue[0].length);
  range.setValues(this.appendQueue);
  this._indices = null;
  this._lastRow = null;
}

// overwrites existing data, be sure to pass everything in
ManagedWorksheet.prototype.update = function(rowIndex, data) {
  if(rowIndex == -1) {
    Logger.log("Asking to update a row with a negative index")
    return;
  }
  var toAppend;
  if(data instanceof Array) {
    toAppend = data;
  } else {
    toAppend = this.objectToRow(data);
  }
  if(this._cache) {
    this._cache[rowIndex - 1] = toAppend;
  } else {
    var range = this.worksheet.getRange(rowIndex, 1, 1, toAppend.length);
    range.setValues([toAppend]);
  }
  this._indices = null;
}

ManagedWorksheet.prototype.activate = function(data) {
  this.worksheet.activate();
}

ManagedWorksheet.prototype.shift = function(data) {
  var row = getRow(2);
  this.removeRow(2);
  return this.rowToObject(row);
}

ManagedWorksheet.prototype.deleteWhere = function(callback) {
  var rowLength = this.getLastRow();
  // start counting at 2 because index of first row is 1, and that first row is the header
  for(var rowIndex=2;rowIndex<=rowLength;rowIndex++) {
    var obj = this.getRowAsObject(rowIndex);
    if(callback.call(this, obj, rowIndex)) {
      this.removeRow(rowIndex);
      rowIndex--;
      rowLength--;
    }
  }
  this._lastRow = null;
  this._indices = null;
}

// [ 30, 'Mary', 'cat'] into { id: 30, name: 'Mary', pet: 'cat' }
ManagedWorksheet.prototype.rowToObject = function(row) {
  if(!row)
    return;
  var obj = {};
  for(var i=0;i<this.columnNames().length;i++) {
    var value = row[i];
    if(typeof(value) === 'undefined') {
      obj[this.columnNames()[i]] = "";
    } else {
      obj[this.columnNames()[i]] = value;
    }
  }
  return obj;
}

// { id: 30, name: 'Mary', pet: 'cat' } into [ 30, 'Mary', 'cat']
ManagedWorksheet.prototype.objectToRow = function(data) {
  var row = [];
  for(var i=0;i<this.columnNames().length;i++) {
    var value = data[this.columnNames()[i]];
    if(typeof(value) === 'undefined') {
      row.push("");
    } else {
      row.push(value);
    }
  }
  return row;
}

ManagedWorksheet.prototype.getRowAsObject = function(rowNumber) {
  var row = this.getRow(rowNumber);
  return this.rowToObject(row);
}

ManagedWorksheet.prototype.getLastColumn = function(rowNumber) {
  if(!this._lastColumn) {
    if(this._cache) {
      this._lastColumn = this._cache[0].length;
    } else {
      this._lastColumn = this.worksheet.getLastColumn();
    }
  }
  return this._lastColumn;
}

ManagedWorksheet.prototype.getLastRow = function(rowNumber) {
  if(!this._lastRow) {
    if(this._cache) {
      this._lastRow = this._cache.length;
    } else {
      this._lastRow = this.worksheet.getLastRow();
    }
  }
  return this._lastRow;
}

ManagedWorksheet.prototype.getRow = function(rowNumber) {
  if(this._cache)
    return this._cache[rowNumber - 1];
  
  var range = this.worksheet.getRange(rowNumber, 1, 1, this.getLastColumn());
  return range.getValues()[0];
}

ManagedWorksheet.prototype.removeRow = function(rowNumber) {
  if(this._cache) {
    this._cache.slice(rowNumber - 1, 1);
  } else {
    this.worksheet.deleteRow(rowNumber);
  }
  this._lastRow = null;
  this._indices = null;
}

ManagedWorksheet.prototype.last = function() {
  var row = this.getLastRow();
  if(row === 1) 
    return;
  return this.getRowAsObject(row);
}

ManagedWorksheet.prototype.getCol = function(colIndex) {
  var values;
  if(this._cache) {
    values = this._cache.map( function(row) { return row[colIndex]; } )
    values.shift();
  } else {
    var range = this.worksheet.getRange(2, colIndex, this.getLastRow() - 1, 1);
    values = range.getValues().map( function(row) { return row[0] } );
  }
  return values;
}

ManagedWorksheet.prototype.indices = function(key, value) {
  if(!this._indices)
    this._indices = {}
  
  if(!this._indices[key]) {
    var colIndex = this.columnNames().indexOf(key);
    this._indices[key] = this.getCol(colIndex + 1);
  }
  
  return this._indices[key];
}

ManagedWorksheet.prototype.allRows = function() {
  if(this._cache)
    return this._cache.slice(1);

  if(this.getLastRow() == 1) {
    return [];
  };
  var range = this.worksheet.getRange(2, 1, this.getLastRow() - 1, this.getLastColumn());
  return range.getValues();
}

ManagedWorksheet.prototype.all = function() {
  var rows = this.allRows();
  
  var that = this;
  return rows.map( function(row) { return that.rowToObject(row) } );
}

ManagedWorksheet.prototype.rowIndex = function(key, value) {
  if(this.getLastRow() == 1)
    return -1;
  
  var index = this.indices(key).indexOf(value);

  if(index === -1)
    return -1;
  
  return index + 2; // need to add 1 because rows start at 1, then another 1 because of headers
}

ManagedWorksheet.prototype.find = function(key, value) {
  var index = this.rowIndex(key, value);
  if(index == -1)
    return;
  var row = this.getRow(index);
  return this.rowToObject(row);
}