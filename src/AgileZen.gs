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
 * Fetches stories from AgileZen.
 * @param {string} The API key that is used to authenticate the request.
 * @param {string} ID of the AgileZen project.
 * @return {object} All available stories; paging is not yet supported.
 */
function fetchStories_(apiKey, projectId) {
  var options = {
    "headers" : {
      "X-Zen-ApiKey" : apiKey
    }
  };

  // https://agilezen.com/api/v1/projects/12345/stories
  var result = UrlFetchApp.fetch(URL_AGILE_ZEN + "/projects/" + projectId + "/stories", options);
  var stories  = Utilities.jsonParse(result.getContentText());
  
  return stories;
};

/**
 * Fetches projects from AgileZen.
 * @param {string} The API key that is used to authenticate the request.
 * @param {string} Name of the AgileZen project.
 * @return {object} All available projects; paging is not yet supported.
 */
function fetchProjects_(apiKey, projectName) {
  var options = {
    "headers" : {
      "X-Zen-ApiKey" : apiKey
    }
  };

  // https://agilezen.com/api/v1/projects/
  var result = UrlFetchApp.fetch(URL_AGILE_ZEN + "/projects/", options);
  var projects  = Utilities.jsonParse(result.getContentText());
  
  return projects;
};

/**
 * Searches the id of an project.
 * @param {string} The API key that is used to authenticate the request.
 * @param {string} Name of the AgileZen project.
 * @return {string} ID of project.
 */
function findProjectIdByName_(apiKey, projectName) {
  var projects = fetchProjects_(apiKey, projectName);
  var projectId = -1;
  if (projectName && projects) {
    for (var i in projects.items) {
      var item = projects.items[i];
      if (item.name.toLowerCase() == projectName.toLowerCase()) {
        projectId = item.id;
        break;
      }
    }
  }
  return projectId;
}

/**
 * Sums up sizes of the given stories, but doesn't consider items that are finished or not ready.
 * @return {number} Sum of sizes of unfinished story items.
 */
function getSize_(stories) {
  // sum of sizes
  var size = 0;
  for (var i in stories.items) {
    var item = stories.items[i];
    if (item.status != AGILE_ZEN_STATE_FINISHED) {
      var itemSize = parseFloat(item.size);
      if (!isNaN(itemSize)) {
        size += parseFloat(item.size);
      }
    }
  }
  return size;
}
