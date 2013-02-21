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
 * The onOpen() function is automatically invoked whenever the spreadsheet is opened.
 */
function onOpen() {
  addMenu_();
};

/**
 * The onInstall() is automatically invoked whenever the spreadsheet is installed.
 */
function onInstall() {
  addMenu_();
};

/**
 * Adds the menu to the spreadsheet.
 */
function addMenu_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : MENU_NEW_BURN_DOWN,
    functionName : "showNewReportDialog"
  }, {
    name : MENU_UPDATE_BURN_DOWN,
    functionName : "showUpdateReportDialog"
  }];
  sheet.addMenu(MENU_ROOT, entries);
}
  

/**
 * Creates a new report.
 * @param {string} The AgileZen API key.
 * @param {number} The time span till the project or iterration ends.
 * @param {number} The time unit.
 * @param {string} The AgileZen project id.
 */
function createReport_(apiKey, time, timeUnit, projectId) {
  var stories = fetchStories_(apiKey, projectId);
  var size = getSize_(stories);
  createSheet_();
  initSheet_(size, time, timeUnit);
  createChart_();
}

/**
 * Retrieves data from AgileZen and updates the sheet.
 * @param {string} The AgileZen API key.
 * @param {number} The time span till the project or iterration ends.
 */
function updateReport_(apiKey, projectId) {
  var stories = fetchStories_(apiKey, projectId);
  var size = getSize_(stories);
  addSizeToSheet_(size);
}

