var DistributionManager = function(options) {
  if(!options)
    options = {}
  this.spreadsheet = new ManagedSpreadsheet(options.key);
  this.document_index = this.spreadsheet.sheet("documents");
  this.worksheet_index = this.spreadsheet.sheet("worksheets");
  this.template = this.spreadsheet.sheet("template");
  
  this.sheetLimit = options.sheetLimit || 200;
  this.baseFolder = options.baseFolder;
  this.folderName = options.folderName || "data";
}

DistributionManager.prototype.documents = function() {
  return this.document_index.all();
}

DistributionManager.prototype.availableSpreadsheet = function() {
  var documents = this.documents();
  var latestKey = documents[documents.length - 1];

  if(!latestKey)
    return this.createDocument();

  var managed = new ManagedSpreadsheet(latestKey.docid);
  
  if(managed.sheetCount > this.sheetLimit)
    return this.createDocument();
  
  return managed;
}

DistributionManager.prototype.createDocument = function() {
  var formattedDate = Utilities.formatDate(new Date(), "EST", "yyyy-MM-dd HH:mm:ss");
  var spreadsheet = SpreadsheetApp.create(formattedDate);
  var managed = new ManagedSpreadsheet(spreadsheet.getId());
  managed.moveToFolder(this.folderName, this.baseFolder);
  this.document_index.append({docid: spreadsheet.getId(), url: spreadsheet.getUrl()});
//  MailApp.sendEmail("user@example.com", "Make this spreadsheet public!", "You've got a new spreadsheet up at " + spreadsheet.getUrl());
  return managed;
}

DistributionManager.prototype.hasKey = function(key) {
  return !!this.worksheet_index.find('key', key);
}

DistributionManager.prototype.template = function(spreadsheet) {
  var template = spreadsheet.getSheetByName("template")
  if(template)
    return template;
  var copied_template = _masterTemplate().copyTo(spreadsheet);
  copied_template.setName("template");
  return copied_template;
  
}

DistributionManager.prototype.getWorksheet = function(key) {
  var worksheet_info = this.worksheet_index.find('key', key);
  if(!worksheet_info)
    return
  
  var managed = new ManagedSpreadsheet(worksheet_info.docid);
  var sheet = managed.sheet(worksheet_info.sheetname);
  return sheet;
}

DistributionManager.prototype.createWorksheet = function(key) {
  if(this.hasKey(key))
    return this.getWorksheet(key);
  
  var managed = this.availableSpreadsheet();
  
  var template = managed.sheet("template");
  if(!template) {
    var template = this.template.copyTo(managed);
  }
  
  var sheet = template.duplicateAs(key);
  this.worksheet_index.append({key: key, sheetid: sheet.getId(), docid: managed.getId(), sheetname: key, url: managed.getUrl() + "#guid=" + sheet.getId() });
  return sheet;
}