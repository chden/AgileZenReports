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
 * Shows form with input fields to create a new report.
 */
function showNewReportDialog() {
  var app = UiApp.createApplication();
  app.setWidth(350);
  app.setHeight(330);
  app.setTitle(TITLE_CREATE_BURN_DOWN);

  // create form
  var grid = app.createGrid(10, 2);
  grid.setWidget(0, 0, app.createLabel(LABEL_PROJECT_NAME));
  grid.setWidget(0, 1, app.createTextBox().setName(PARAMETER_PROJECT_NAME));
  grid.setWidget(1, 1, app.createLabel(LABEL_HELP_PROJECT_NAME).setStyleAttribute("fontSize","10px"));
  grid.setWidget(2, 0, app.createLabel(LABEL_API_KEY));
  grid.setWidget(2, 1, app.createTextBox().setName(PARAMETER_API_KEY).setValue(loadApiKey_()));
  grid.setWidget(3, 1, app.createLabel(LABEL_HELP_API_KEY).setStyleAttribute("fontSize","10px"));
  grid.setWidget(5, 0, app.createLabel(LABEL_REMAINING_TIME));
  var textBoxTime = app.createTextBox().setName(PARAMETER_REMAINING_TIME);
  grid.setWidget(5, 1, textBoxTime);
  grid.setWidget(6, 1, app.createLabel(LABEL_HELP_REMAINING_TIME).setStyleAttribute("fontSize","10px"));
  grid.setWidget(7, 0, app.createLabel(LABEL_REMAINING_TIME_UNIT));
  var textBoxTimeUnit = app.createTextBox().setName(PARAMETER_TIME_UNIT);
  grid.setWidget(7, 1, textBoxTimeUnit);
  grid.setWidget(8, 1, app.createLabel(LABEL_HELP_TIME_UNIT).setStyleAttribute("fontSize","10px"));

  var grid2 = app.createGrid(4, 1);
  var buttonOk = app.createButton(BUTTON_CREATE);
  grid2.setWidget(0, 0, buttonOk);  
  var errorMessage1 = app.createLabel(LABEL_ERROR_MESSAGE1).setVisible(false);
  grid2.setWidget(2, 0, errorMessage1);
  var errorMessage2 = app.createLabel(LABEL_ERROR_MESSAGE2).setVisible(false);
  grid2.setWidget(3, 0, errorMessage2);

  // validate input
  buttonOk.addClickHandler(app.createClientHandler()
      .validateNotRange(textBoxTime, 1, null)
      .forTargets(buttonOk).setEnabled(false)
      .forTargets(errorMessage1).setVisible(true));
  textBoxTime.addClickHandler(app.createClientHandler()
      .forTargets(buttonOk).setEnabled(true)
      .forTargets(errorMessage1).setVisible(false));
  buttonOk.addClickHandler(app.createClientHandler()
      .validateNotRange(textBoxTimeUnit, 1, null)
      .forTargets(buttonOk).setEnabled(false)
      .forTargets(errorMessage2).setVisible(true));
  textBoxTimeUnit.addClickHandler(app.createClientHandler()
      .forTargets(buttonOk).setEnabled(true)
      .forTargets(errorMessage2).setVisible(false));
  buttonOk.addClickHandler(app.createServerHandler("createReportHandler_")
      .addCallbackElement(grid)
      .validateRange(textBoxTime, 1, null)
      .validateRange(textBoxTimeUnit, 1, null));

  app.add(grid);
  app.add(grid2);

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  spreadsheet.show(app);
}

/**
 * Handles the input of the form fields to create a new report.
 */
function createReportHandler_(e) {
  var apiKey = e.parameter.apiKey;
  storeApiKey_(apiKey);
  var time = parseFloat(e.parameter.time);
  var timeUnit = parseFloat(e.parameter.timeUnit);
  var projectName = e.parameter.projectName;
  var projectId = findProjectIdByName_(apiKey, projectName);
  createReport_(apiKey, time, timeUnit, projectId);
  
  storeProjectName_(projectName);
  storeProjectId_(projectId);

  var app = UiApp.getActiveApplication();
  app.close();
  return app;
}

/**
 * Shows form with input fields to update a new report.
 */
function showUpdateReportDialog() {
  var app = UiApp.createApplication();
  app.setWidth(200);
  app.setHeight(90);
  app.setTitle(TITLE_UPDATE_BURN_DOWN);

  // create form
  var grid = app.createGrid(2, 2);
  grid.setWidget(0, 0, app.createLabel(LABEL_API_KEY));
  grid.setWidget(0, 1, app.createTextBox().setName(PARAMETER_API_KEY).setValue(loadApiKey_()));
  var buttonOk = app.createButton(BUTTON_UPDATE);  
  buttonOk.addClickHandler(app.createServerHandler("updateReportHandler_")
      .addCallbackElement(grid));

  app.add(grid);
  app.add(buttonOk);

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  spreadsheet.show(app);
}

/**
 * Handles the input of the form fields to update a report.
 */
function updateReportHandler_(e) {
  var apiKey = e.parameter.apiKey;
  storeApiKey_(apiKey);
  var projectId = loadProjectId_();
  updateReport_(apiKey, projectId);

  var app = UiApp.getActiveApplication();
  app.close();
  return app;
}

/**
 * Inserts a new chart in the active sheet.
 */
function createChart_() {
  var sheet = SpreadsheetApp.getActiveSheet();
  // create a new chart with some options
  // see https://google-developers.appspot.com/chart/interactive/docs/gallery/linechart#Configuration_Options
  var chartBuilder = sheet.newChart()
    .setPosition(6, 4, 5, 5)
    .setChartType(Charts.ChartType.LINE)
    .addRange(sheet.getRange("A:C"))
    .setOption("title", CHART_TITLE)
    .setOption("hAxis.title", CHART_HAXIS_TITLE)
    .setOption("vAxis.title", CHART_VAXIS_TITLE)
    .setOption("legend.position", CHART_LEGEND_POSITION)
    .setOption("interpolateNulls", false)
    .setOption("pointSize", 5);
  
  sheet.insertChart(chartBuilder.build());
}
