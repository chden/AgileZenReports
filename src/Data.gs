/**
 * @preserve
 * Copyright 2012 Christoph Denninger
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */


/**
 * Loads the user specific AgileZen API key from the user properties service.
 * @return {string} The AgileZen API key, or null if the key hasn't been stored yet.
 */
function loadApiKey_() {
  return UserProperties.getProperty(USER_PROPERTY_AGILE_ZEN_API_KEY);
}

/**
 * Stores the Agile Zen API key in the user properties service.
 * @param {string} The AgileZen API key.
 */
function storeApiKey_(apiKey) {
  UserProperties.setProperty(USER_PROPERTY_AGILE_ZEN_API_KEY, apiKey);
}

/**
 * Loads the project name from the active sheet.
 * @return {string} The AgileZen project name, or null if the name hasn't been stored yet.
 */
function loadProjectName_(projectName) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var value = sheet.getRange("F2");
  return value.getValue();
}

/**
 * Stores the name of the Agile Zen project in the active sheet.
 * @param {string} The AgileZen project name.
 */
function storeProjectName_(projectName) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var key = sheet.getRange("E2");
  key.setValue(SHEET_CELL_PROJECT_NAME)
      .setBackground(BACKGROUND_COLOR_GREY)
      .setFontWeight(FONT_WEIGHT_BOLD)
      .setHorizontalAlignment(ALIGNMENT_LEFT)
      .setBorder(true, true, false, false, false, false);
  var value = sheet.getRange("F2");
  value.setValue(projectName)
      .setBackground(BACKGROUND_COLOR_GREY)
      .setHorizontalAlignment(ALIGNMENT_LEFT)
      .setBorder(true, false, false, true, false, false);
}

/**
 * Loads the project id from the active sheet.
 * @return {string} The AgileZen project id, or null if the name hasn't been stored yet.
 */
function loadProjectId_() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var value = sheet.getRange("F3");
  return value.getValue();
}

/**
 * Stores the id of the Agile Zen project in the active sheet.
 * @param {string} The AgileZen project id.
 */
function storeProjectId_(projectId) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var key = sheet.getRange("E3");
  key.setValue(SHEET_CELL_PROJECT_ID)
      .setBackground(BACKGROUND_COLOR_GREY)
      .setFontWeight(FONT_WEIGHT_BOLD)
      .setHorizontalAlignment(ALIGNMENT_LEFT)
      .setBorder(false, true, true, false, false, false);
  var value = sheet.getRange("F3");
  value.setValue(projectId)
      .setBackground(BACKGROUND_COLOR_GREY)
      .setHorizontalAlignment(ALIGNMENT_LEFT)
      .setBorder(false, false, true, true, false, false);
}

/**
 * Fills the sheet with values.
 * @param {number} Sum of sizes of unfinished items in the backlog.
 * @param {number} The overall time span, same unit as size, till the project or iterration ends.
 * @param {number} The time unit.
 */
function initSheet_(size, timeOverall, timeUnit) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var cell = sheet.getRange("A1");
  
  // overall time is increased by one to include the initial backlog state
  time = timeOverall + 1;

  // set column headers
  cell.offset(0, 0).setValue(SHEET_CELL_ITERATION_TIMELINE)
      .setBackground(BACKGROUND_COLOR_GREEN)
      .setFontWeight(FONT_WEIGHT_BOLD)
      .setHorizontalAlignment(ALIGNMENT_CENTER)
      .setBorder(false, false, true, false, false, false);
  sheet.setColumnWidth(1, 155);
  cell.offset(0, 1).setValue(SHEET_CELL_IDEAL_REMAINING_EFFORT)
      .setBackground(BACKGROUND_COLOR_GREEN)
      .setFontWeight(FONT_WEIGHT_BOLD)
      .setHorizontalAlignment(ALIGNMENT_CENTER)
      .setBorder(false, false, true, false, false, false);
  sheet.setColumnWidth(2, 155);
  cell.offset(0, 2).setValue(SHEET_CELL_ACTUAL_REMAINING_EFFORT)
      .setBackground(BACKGROUND_COLOR_GREEN)
      .setFontWeight(FONT_WEIGHT_BOLD)
      .setHorizontalAlignment(ALIGNMENT_CENTER)
      .setBorder(false, false, true, false, false, false);
  sheet.setColumnWidth(3, 155);

  // insert some data
  cell.offset(1, 0).setFormula('TO_TEXT(0)');
  cell.offset(1, 1).setFormula('C2');
  for (i = 1; i < time; i++) {
    var ii = i + 1;
    var iii = i + 2;
    cell.offset(ii, 0).setFormula('TO_TEXT(' + i + ' * ' + timeUnit + ')');
    //cell.offset(ii, 1).setFormula('C2 - (C2 / ((COUNTIF(A:A, ">-1") - 1) * (A' + iii + ' - A' + ii + '))) * A' + iii);
    cell.offset(ii, 1).setFormula('C2 - (C2 / ((COUNTA(A:A) - 2) * (A' + iii + ' - A' + ii + '))) * A' + iii);
  }
  
  // align first column
  sheet.getRange("A:A").setHorizontalAlignment("right");
  
  // set a number format of the second column
  // sheet.getRange("B1:B").setNumberFormat("#.##");

  // set initial size
  cell.offset(1, 2).setValue(size);
}

/**
 * Updates the active sheet by adding the current size of the backlog to "Actual Remaining Effort" column.
 * @param {number} Sum of sizes of unfinished items in the backlog.
 */
function addSizeToSheet_(size) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var firstEmptyRow = getFirstEmptyRow_(sheet.getRange("C:C"));
  sheet.getRange(firstEmptyRow, 3).setValue(size);
}
