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
 * Inserts a new sheet into the active spreadsheet.
 */
function createSheet_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  //sheet.insertSheet(name + " " + sheet.getNumSheets());
  sheet.insertSheet();
}

/**
 * Calculates the first empty row of a given range.
 * @param {List.<Range>} The range to search for the empty row.
 * @return {number} index of first empty row in range.
 */
function getFirstEmptyRow_(range) {
  var values = range.getValues();
  var counter = 0;
  while (values[counter][0] != "") {
    counter++;
  }
  return (counter + 1);
}
