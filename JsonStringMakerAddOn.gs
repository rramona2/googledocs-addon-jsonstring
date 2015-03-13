/**
 * This is free and unencumbered software released into the public domain.
 *
 * Anyone is free to copy, modify, publish, use, compile, sell, or
 * distribute this software, either in source code form or as a compiled
 * binary, for any purpose, commercial or non-commercial, and by any
 * means.
 *
 * In jurisdictions that recognize copyright laws, the author or authors
 * of this software dedicate any and all copyright interest in the
 * software to the public domain. We make this dedication for the benefit
 * of the public at large and to the detriment of our heirs and
 * successors. We intend this dedication to be an overt act of
 * relinquishment in perpetuity of all present and future rights to this
 * software under copyright law.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
 * EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
 * MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
 * IN NO EVENT SHALL THE AUTHORS BE LIABLE FOR ANY CLAIM, DAMAGES OR
 * OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE,
 * ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
 * OTHER DEALINGS IN THE SOFTWARE.
 *
 * For more information, please refer to <http://unlicense.org/>
 */

/**
 * JsonStringMakerAddOn.gs
 * Designed for Google Drive Spreadsheeets.
 * Converts Spreadsheet range of selected values into JSON array strings described here:
 *   http://www.json.org/
 *
 * @license The Unlicense http://unlicense.org/
 * @version 0.8
 * @updated 2015-03-12
 * @author  The Pffy Authors https://github.com/pffy/
 * @link    https://github.com/pffy/googledocs-addon-jsonstring
 *
 */
var product = {
  
  "name": "JsonStringMaker Add-On",
  "version": "0.8",
  
  "license": "This is free, libre and open source software.",
  "licenseUrl": "http://unlicense.org/",
  
  "author": "The Pffy Authors",
  "authorUrl": "https://github.com/pffy",
  
  "sidebarTitle": "Selected JSON String",
  "sidebarFilename": "JsonStringPreview",
 
  "exportJsonFilename": "NewJsonString.json", 
  
  "tagline": "Live long and prosper."
  
}

var menuItems = {
 
  "about": "About " + product.name,
  "help": product.name + " Help Guide",
  
  "convert": "Convert range to JSON array string ...",
  "save": "Save range as JSON array in file ..."
  
}

var messages = {
  
  "saved": "saved to Google Drive.",
  "done": "task completed."
  
}

var tips = {
 
  "selectTip": "\n\nSELECT spreadsheet rows and columns of cells to create a range.",
  "convertTip": "\n\nCONVERT that range to a JSON string to be copied and pasted.",
  
  "saveJsonTip": "\n-OR-\nSAVE it as " + product.exportJsonFilename 
    + " on Google Drive."  
  
}


/**
 * Add-On UI/Menus
 **/


// converts range to json
function _convertJson() {
  var range = SpreadsheetApp.getActiveSheet().getActiveRange();
  return '' + new JsonStringMaker(range);
}

// saves range as JSON to new Google Drive document with JSON file extension
function _saveAsJson(filename) {

  var range = SpreadsheetApp.getActiveSheet().getActiveRange();
  var mtm = new JsonStringMaker(range);

  DriveApp.createFile(filename, '' + mtm, MimeType.PLAIN_TEXT);
  SpreadsheetApp.getUi()
     .alert(filename + ' ' + messages.saved);
}

// handles convert menu item
function _convertItem() {

  var ui, jsonString;

  ui = HtmlService
  .createHtmlOutputFromFile(product.sidebarFilename)
  .setTitle(product.sidebarTitle);

  jsonString = _convertJson();
  ui.append('<br/>Copy and Paste this:<br/><textarea>' + jsonString + '</textarea>');

  SpreadsheetApp.getUi().showSidebar(ui);
}

// handles save menu item
function _saveItem() {
  _saveAsJson(product.exportJsonFilename);
}

// handles help menu item
function _helpItem() {
  SpreadsheetApp.getUi()
     .alert(_getHelpInfo());
}

// handles about menu item
function _aboutItem() {
  SpreadsheetApp.getUi()
     .alert(_getAboutInfo());
}

// returns Help information
function _getHelpInfo() {

  var str = '';
  str += menuItems.help;
  
  str += tips.selectTip;
  str += tips.convertTip;
  str += tips.saveJsonTip;
  
  str += '\n\n' + product.tagline;
  str += '\n --' + product.author;

  return str;
}

// returns About information: product, license and authors
function _getAboutInfo() {

  var str = '';
  
  str += product.name;
  str += '\nVersion ' + product.version;

  str += '\n\n' + product.license;
  str += '\n' + product.licenseUrl;

  str += '\n\n' + product.author;
  str += '\n' + product.authorUrl;
  
  return str;
}

// adds menus to Google Drive Spreadsheets
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu(product.name)
      .addItem(menuItems.convert, '_convertItem')
      .addItem(menuItems.save, '_saveItem')
      .addSeparator()
      .addItem(menuItems.help, '_helpItem')
      .addSeparator()
      .addItem(menuItems.about, '_aboutItem')
      .addToUi();
}


