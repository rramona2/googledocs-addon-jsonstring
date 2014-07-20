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
 * @version 0.1
 * @updated 2014-07-20
 * @author  The Pffy Authors https://github.com/pffy/
 * @link    https://github.com/pffy/googledocs-addon-jsonstring
 *
 */

// get product name
function getProductName() {
  return 'JsonStringMaker';
}

// get product title
function getSidebarTitle() {
  return 'Selected JSON String';
}

function getExportFilename() {
  return 'NewJsonString.js';
}

// converts range to json
function convertJson() {
  var range = SpreadsheetApp.getActiveSheet().getActiveRange();
  return '' + new JsonStringMaker(range);
}

// saves range as JSON to new Google Drive document with JSON file extension
function saveAsJson(filename) {

  var range = SpreadsheetApp.getActiveSheet().getActiveRange();
  var mtm = new JsonStringMaker(range);

  DriveApp.createFile(filename, '' + mtm, MimeType.PLAIN_TEXT);
  SpreadsheetApp.getUi()
     .alert(filename + ' saved to Google Drive.');
}

// handles convert menu item
function convertItem() {

  var ui, jsonString;

  ui = HtmlService
  .createHtmlOutputFromFile('JsonStringPreview')
  .setTitle(getSidebarTitle());

  jsonString = convertJson();
  ui.append('<br/>Copy and Paste this:<br/><textarea>' + jsonString + '</textarea>');

  SpreadsheetApp.getUi().showSidebar(ui);
}

// handles save menu item
function saveItem() {
  saveAsJson(getExportFilename());
}

// handles help menu item
function helpItem() {
  SpreadsheetApp.getUi()
     .alert(getHelpInfo());
}

// handles about menu item
function aboutItem() {
  SpreadsheetApp.getUi()
     .alert(getAboutInfo());
}

// returns Help information
function getHelpInfo() {

  var str = '';

  str += 'HELP GUIDE';
  str += '\n\nSELECT spreadsheet rows and columns of cells to create a range.';
  str += '\n\nCONVERT that range to JSON to be copied and pasted.';
  str += '\n-OR-\nSAVE it as ' + getExportFilename()
    + ' on Google Drive, to edit with Drive apps.';

  str += '\n\nLive long and prosper.';
  str += '\n --The Pffy Authors';

  return str;
}

// returns About information: product, license and authors
function getAboutInfo() {

  var str = '';

  str += getProductName();
  str += '\nVersion 0.1';

  str += '\n\nThis is free, libre and open source software.';
  str += '\nhttp://unlicense.org/';

  str += '\n\nThe Pffy Authors';
  str += '\nhttps://github.com/pffy';

  return str;
}

// adds menus to Google Drive Spreadsheets
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu(getProductName())
      .addItem('Convert range to JSON array string ...', 'convertItem')
      .addItem('Save range as JSON array in file ...', 'saveItem')
      .addSeparator()
      .addItem('Help Guide', 'helpItem')
      .addSeparator()
      .addItem('About ' + getProductName(), 'aboutItem')
      .addToUi();
}


