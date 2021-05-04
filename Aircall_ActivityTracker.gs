/*
* Main function that controls the logic of the Activity Tracker
*/
function activityTracker() {
  /*
  These are user-specified constants that the user must set beforehand which will
  determine the overall behaviour of the Activity Tracker.
  */
  // Define how far back (in mins) the feed should review Aircall phonecall history
  var INTERVAL = 15;
  // Define how often an Aircall Calls API call should be made to refresh Aircall user information
  var USER_REFRESH_INTERVAL = 240;
  // Define how often an Aircall Teams API call should be made to refresh Aircall teams information
  var TEAM_REFRESH_INTERVAL = 480;
  // Define how often an Aircall Numbers API call should be made to refresh Aircall numbers information
  var NUM_REFRESH_INTERVAL = 480;
  // Define the threshold of what is considered an "acceptable" time for a call to be waiting
  // in queue (i.e. Acceptable < WAIT_LEVEL_1 < Caution/Warning < WAIT_LEVEL_2 < Unacceptable)
  var WAIT_LEVEL_1 = 10;
  var WAIT_LEVEL_2 = 20;
  // Define the acceptable service level (e.g. avg. amount of time a caller waits before agent picks
  // up the phone) for each user. If "0", this feature is disabled.
  var SLA = 0;
  // The API key that will be used to make API calls to Aircall openAPI
  var API_KEY;

  /*
  These are constants that should not be modified.
  */
  // Need to define special ID for users with no Aircall team
  const NO_TEAM = "999999999"
  // Define cell sizes for different cell types
  const WAIT_CELL_ROWSIZE = 3;
  const WAIT_CELL_COLSIZE = 2;
  const DIVIDER_ROWSIZE = 1;
  const TEAM_CELL_ROWSIZE = 4;
  const TEAM_CELL_COLSIZE = 2;
  const USER_CELL_ROWSIZE = 2;
  const USER_CELL_COLSIZE = 2;
  const NUM_SUMMARY_ROWSIZE = 1;
  const NUM_SUMMARY_COLSIZE = 2;


  // Write all changes to the Google spreadsheet
  SpreadsheetApp.flush();

  // Retrieve all the metadata that has been cached (for performance reasons)
  var cache = PropertiesService.getScriptProperties();

  // We will override the configuration parameters set in the script with those input from
  // the configuration params sidebar
  if (cache.getProperty("interval")) {INTERVAL = parseInt(cache.getProperty("interval")).toFixed(0);}
  if (cache.getProperty("api_key")) {API_KEY = cache.getProperty("api_key");}
  if (cache.getProperty("user_interval")) {USER_REFRESH_INTERVAL = cache.getProperty("user_interval");}
  if (cache.getProperty("team_interval")) {TEAM_REFRESH_INTERVAL = cache.getProperty("team_interval");}
  if (cache.getProperty("wait_lvl_1")) {
    WAIT_LEVEL_1 = cache.getProperty("wait_lvl_1");
  } else {
    cache.setProperty("wait_lvl_1", WAIT_LEVEL_1);
  }
  if (cache.getProperty("wait_lvl_2")) {
    WAIT_LEVEL_2 = cache.getProperty("wait_lvl_2");
  } else {
    cache.setProperty("wait_lvl_2", WAIT_LEVEL_2);
  }
  if (cache.getProperty("sla")) {SLA = parseInt(cache.getProperty("sla"));}


  // If we notice we are using a new API key (most likely to pull info from a new
  // Aircall instance), we will force a refresh of Aircall user and team info
  if (parseInt(cache.getProperty("api_key_changed"))) {
    Logger.log("New API key set, wiping the cache...")
    cache.setProperty("api_key", API_KEY);
    cache.setProperty("api_key_changed", "0");
    cache.deleteProperty("user_cache_time");
    cache.deleteProperty("team_cache_time");
  }

  // If no API key is still set, exit the program
  if (!API_KEY) {
    SpreadsheetApp.getUi().alert("No API key set, will not continue.");
    return;
  }

  // Set headers and parameters for any future Aircall API calls
  var headers = {
    "Authorization" : "Basic " + API_KEY
  };
  var get_params = {
    "method":"GET",
    "headers":headers
  };

  // Create a JSON object which will contain the general structure of the Activity Tracker
  // within the spreadsheet
  // Structure: {
  //  "Team ID 1": [List of Users part of Team 1],
  //  "Team ID 2": [List of Users part of Team 2],
  //  ...
  // }
  var livefeed_results = {};


  // Used for pagination when making calls to Aircall API
  var fetch_nextpage = true;

  // Determine current time
  var user_ref_check_time = parseInt((new Date().getTime())/1000);

  // Determine last time Aircall user info was refreshed
  var user_ref_cache_time;
  if (cache.getProperty("user_cache_time")) {
    user_ref_cache_time = parseInt(cache.getProperty("user_cache_time"));
  } else {
    user_ref_cache_time = user_ref_check_time - ((USER_REFRESH_INTERVAL+1) * 60);
  }

  // If Aircall user info is due for refreshing (cache expiry period complete)...
  if ((user_ref_check_time - user_ref_cache_time) > (USER_REFRESH_INTERVAL * 60)) {
    Logger.log("Refreshing IDs of Aircall users...");

    // Get the "UserDB" sheet
    var userdb_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UserDB");
    userdb_sheet.getRange("A:B").clearContent();

    // Create a data structure (array) to store User ID and Username
    // Structure: [
    //  ["User ID 1", "Username 1"],
    //  ["User ID 2", "Username 2"],
    //  ...
    // ]
    var userdb_results = [];

    // Define the url for the Aircall User API call
    var users_url = "https://api.aircall.io/v1/users?per_page=50";

    // While there is still at least one page to review for Aircall user information
    // (pagination)...
    while (fetch_nextpage) {

      // Make the Aircall User API call and parse the response
      var users_response = UrlFetchApp.fetch(users_url, get_params);
      var users_json = users_response.getContentText();
      var users_data = JSON.parse(users_json);

      // For each Aircall user...
      for (var i = 0; i < users_data.users.length; i++) {
        // Keep track of the user's user ID and username
        userdb_results.push([users_data.users[i].id.toString(), users_data.users[i].name]);
      }

      // Determine if we need to iterate through other pages of results from the API call...
      if (users_data.meta.next_page_link) {
        users_url = users_data.meta.next_page_link;
      } else {
        fetch_nextpage = false;
      }
    }

    // Write the Aircall user information to the "UserDB" sheet
    userdb_sheet.getRange("A1:B" + userdb_results.length.toString()).setValues(userdb_results);

    // Keep track of the refresh time
    cache.setProperty("user_cache_time", user_ref_check_time.toString());
  }


  // Create a JSON object which will contain the mappings between User ID and Username
  // Structure: {
  //  "User ID 1": "Username 1",
  //  "User ID 2": "Username 2",
  //  ...
  // }
  var userdb = {};

  // Create a JSON object which will contain the mappings between User ID and Phone Numbers
  // Structure: {
  //  "User ID 1": {
  //    "Phone Number ID A": "Phone Number Name A",
  //    "Phone Number ID B": "Phone Number Name B",
  //    ...
  //   },
  //  "User ID 2": {
  //    "Phone Number ID G": "Phone Number Name G",
  //    "Phone Number ID H": "Phone Number Name H",
  //    ...
  //   },
  //  ...
  // }
  var user_numberdb = {};


  // From the "UserDB" sheet, collect all the Aircall user info to keep track of mappings
  // between User ID and Username
  userdb_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UserDB");
  var userdb_data = userdb_sheet.getDataRange().getValues();
  for (var i = 0; i < userdb_data.length; i++) {
    userdb[userdb_data[i][0].toString()] = userdb_data[i][1];

    // If we know which numbers this user is assigned to...
    if (userdb_data[i][2]) {
      // Convert the string into a list of phone numbers (delimited by the ";" character)
      var numbers = userdb_data[i][2].split(";");

      user_numberdb[userdb_data[i][0].toString()] = {};

      // For each phone number...
      for (var n = 0; n < numbers.length; n++) {
        // Keep track of it
        user_numberdb[userdb_data[i][0].toString()][numbers[n].split(":")[0]] = numbers[n].split(":")[1];
      }
    }
  }

  // Determine if users have been mapped to numbers (which will enable certain features)
  var user_num_mappings_enabled = false;
  if (Object.keys(user_numberdb).length > 0) {
    user_num_mappings_enabled = true;
  }

  // Used for pagination when making calls to Aircall API
  var fetch_nextpage = true;

  // Determine current time
  var number_ref_check_time = parseInt((new Date().getTime())/1000);

  // Determine last time Aircall number info was refreshed
  var number_ref_cache_time;
  if (cache.getProperty("number_cache_time")) {
    number_ref_cache_time = parseInt(cache.getProperty("number_cache_time"));
  } else {
    number_ref_cache_time = number_ref_check_time - ((NUM_REFRESH_INTERVAL+1) * 60);
  }

  // If Aircall number info is due for refreshing (cache expiry period complete)...
  if ((number_ref_check_time - number_ref_cache_time) > (NUM_REFRESH_INTERVAL * 60)) {
    Logger.log("Refreshing IDs of Aircall numbers...");

    // Get the "NumberDB" sheet
    var numberdb_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NumberDB");
    numberdb_sheet.clear();

    // Create a data structure (array) to store Number ID and Number name
    // Structure: [
    //  ["Number ID 1", "Number Name 1"],
    //  ["Number ID 2", "Number Name 2"],
    //  ...
    // ]
    var numberdb_results = [];
    numberdb_sheet.clear();

    // Define the url for the Aircall Number API call
    var numbers_url = "https://api.aircall.io/v1/numbers?per_page=50";

    // While there is still at least one page to review for Aircall number information
    // (pagination)...
    while (fetch_nextpage) {

      // Make the Aircall User API call and parse the response
      var numbers_response = UrlFetchApp.fetch(numbers_url, get_params);
      var numbers_json = numbers_response.getContentText();
      var numbers_data = JSON.parse(numbers_json);

      // For each Aircall number...
      for (var i = 0; i < numbers_data.numbers.length; i++) {
        // Keep track of the number's ID and name
        numberdb_results.push([numbers_data.numbers[i].id.toString(), numbers_data.numbers[i].name]);
      }

      // Determine if we need to iterate through other pages of results from the API call...
      if (numbers_data.meta.next_page_link) {
        numbers_url = numbers_data.meta.next_page_link;
      } else {
        fetch_nextpage = false;
      }
    }

    // Write the Aircall number information to the "NumberDB" sheet
    numberdb_sheet.getRange("A1:B" + numberdb_results.length.toString()).setValues(numberdb_results);

    // Keep track of the refresh time
    cache.setProperty("number_cache_time", number_ref_check_time.toString());
  }

  // Check if there is a filter on the Aircall teams we want to view
  // Structure: {
  //    "Team ID 1": true,
  //    "Team ID 2": true,
  //    ...
  // }
  var filtered_teams = JSON.parse(cache.getProperty("filteredteams_cache_db"));

  // If either there are no teams we want to filter on, or there is a team filter
  // and it contains the psuedo-group of single users...
  if (!filtered_teams || filtered_teams[NO_TEAM]) {
    livefeed_results[NO_TEAM] = [];
  }


  // Get the "Activity Tracker" sheet (the main page)
  var livefeed_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Activity Tracker");

  // Create a JSON object which will contain the mappings between Aircall
  // Team IDs and Team Names
  // Structure: {
  //  "Team ID 1": "Team Name 1",
  //  "Team ID 2": "Team Name 2",
  //  ...
  // }
  var teamnamesdb = {};
  teamnamesdb[NO_TEAM] = "Single Users";

  // Create a JSON object which will keep track of which Aircall users belong
  // to which Aircall teams
  // Structure: {
  //  "User ID 1": {
  //    "Team ID A": "Team Name A",
  //    "Team ID B": "Team Name B",
  //    ...
  //  },
  //  "User ID 2": {
  //    "Team ID C": "Team Name C",
  //    "Team ID D": "Team Name D",
  //    ...
  //  }
  //  ...
  // }
  var teamsdb = {};

  // Determine current time
  var team_ref_check_time = parseInt((new Date().getTime())/1000);

  // Determine last time Aircall team info was refreshed
  var team_ref_cache_time;
  if (cache.getProperty("team_cache_time")) {
    team_ref_cache_time = parseInt(cache.getProperty("team_cache_time"));
  } else {
    team_ref_cache_time = team_ref_check_time - ((TEAM_REFRESH_INTERVAL+1) * 60);
  }


  // If the team filter has changed recently...
  if (parseInt(cache.getProperty("filteredteams_changed"))) {
    // Force refresh the user-to-team mappings
    team_ref_cache_time = team_ref_check_time - ((TEAM_REFRESH_INTERVAL+1) * 60);
    cache.setProperty("filteredteams_changed", "0");
  }


  // If Aircall team info is due for refreshing (cache expiry period complete)...
  if ((team_ref_check_time - team_ref_cache_time) > (TEAM_REFRESH_INTERVAL * 60)) {
    Logger.log("Refreshing Aircall teams...");

    fetch_nextpage = true;

    // Define the url for the Aircall Team API call
    var teams_url = "https://api.aircall.io/v1/teams?per_page=50";

    // While there is still at least one page to review for Aircall team information
    // (pagination)...
    while (fetch_nextpage) {

      // Make the Aircall Team API call and parse the response
      var teams_response = UrlFetchApp.fetch(teams_url, get_params);
      var teams_json = teams_response.getContentText();
      var teams_data = JSON.parse(teams_json);

      // For every Aircall Team...
      for (var i = 0; i < teams_data.teams.length; i++) {

        // If there is no team filter, or if there is a team filter and this
        // team is part of the filter...
        if (!filtered_teams ||
        (filtered_teams && filtered_teams[teams_data.teams[i].id])) {
          // Start making the headers of the Activity Tracker
          livefeed_results[teams_data.teams[i].id] = [];

          // Keep track of the mapping between Team ID and Team Name
          teamnamesdb[teams_data.teams[i].id.toString()] = teams_data.teams[i].name;
        }

        // For every Aircall user that belongs to this team...
        for (var j = 0; j < teams_data.teams[i].users.length; j++) {
          var teamuser = teams_data.teams[i].users[j].id.toString();

          // Add the current team to the list of teams this Aircall user is part of
          if (!teamsdb[teamuser]) {
            teamsdb[teamuser] = {};
            teamsdb[teamuser][teams_data.teams[i].id.toString()] = teams_data.teams[i].name;
          } else {
            teamsdb[teamuser][teams_data.teams[i].id.toString()] = teams_data.teams[i].name;
          }
        }
      }

      // Determine if we need to iterate through other pages of results from the API call...
      if (teams_data.meta.next_page_link) {
        teams_url = teams_data.meta.next_page_link;
      } else {
        fetch_nextpage = false;
      }
    }

    // Write the Aircall team information (mapping between Team IDs and Team Names, as well as
    // mapping of users to teams) to the cache for performance reasons
    cache.setProperty("teams_cache_db", JSON.stringify(teamsdb));
    cache.setProperty("teamnames_cache_db", JSON.stringify(teamnamesdb));

    // Keep track of refresh time
    cache.setProperty("team_cache_time", team_ref_check_time.toString());

  // ... else, we should utilize the cache to skip the Aircall API Teams call...
  } else {
    teamnamesdb = JSON.parse(cache.getProperty("teamnames_cache_db"));
    teamsdb = JSON.parse(cache.getProperty("teams_cache_db"));

    // Start making the headers of the Activity Tracker
    var all_teamnames = Object.keys(teamnamesdb);

    // For each Aircall Team...
    for (var i = 0; i < all_teamnames.length; i++) {

      // If there is no team filter, or if there is a team filter and this
      // team is part of the filter...
      if (!filtered_teams || (filtered_teams && filtered_teams[all_teamnames[i]])) {
        livefeed_results[all_teamnames[i]] = [];
      }
    }
  }

  Logger.log("Queried for all Aircall teams...");


  // Create a JSON object which will keep track of which Aircall users are currently on a call
  // Structure: {
  //  "User ID 1": {
  //    "direction": <direction>,
  //    "number": <external_phonenum>,
  //    "aircall_number":  <participating_aircall_phonenum>,
  //    "calltime": <current_length_of_call>
  //  },
  //  "User ID 2": {
  //    "direction": <direction>,
  //    "number": <external_phonenum>,
  //    "aircall_number":  <participating_aircall_phonenum>,
  //    "calltime": <current_length_of_call>
  //  }
  //  ...
  // }
  var callsdb = {};

  // Create a JSON object which will keep track of all inbound calls which have not
  // yet been answered by an Aircall user
  // Structure: {
  //  "Number ID 1": {
  //    "aircall_number":  <participating_aircall_phonenum>,
  //    "number": <external_phonenum>,
  //    "waittime": <current_waiting_time>
  //    "aircall_number_id": <aircall_phonenum_id>
  //    "aircall_number_users": <list_of_users_assigned_to_number>
  //    "color": <border_color_of_cell>
  //    "livefeed_cells": <list_of_cells_of_users_assigned_to_number>
  //  },
  //  "Number ID 2": {
  //    "aircall_number":  <participating_aircall_phonenum>,
  //    "number": <external_phonenum>,
  //    "waittime": <current_waiting_time>
  //    "aircall_number_id": <aircall_phonenum_id>
  //    "aircall_number_users": <list_of_users_assigned_to_number>
  //    "color": <border_color_of_cell>
  //    "livefeed_cells": <list_of_cells_of_users_assigned_to_number>
  //  }
  //  ...
  // }
  var callswaitingdb = {};

  // Create a JSON object which will keep track summary of all inbound calls and
  // outbound calls for each Aircall user
  // Structure: {
  //  "User ID 1": [<number_of_inbound_calls>, <number_of_outbound_calls>],
  //  "User ID 2": [<number_of_inbound_calls>, <number_of_outbound_calls>],
  //  ...
  // }
  var numcallsdb = {};

  var avail_pagenum = 1;
  fetch_nextpage = true;

  // Convert current time and [current time - Activity Tracker interval period] to
  // UNIX epoch format
  var currtime = parseInt((new Date().getTime())/1000);
  var mins_before_currtime = parseInt(currtime - (INTERVAL * 60));

  // Define the url for the Aircall Calls API call
  var calls_url = "https://api.aircall.io/v1/calls?order=desc&per_page=50&from=" + mins_before_currtime.toString() +
  "&to=" + currtime.toString();

  // While there is still at least one page to review for Aircall call information
  // (pagination)...
  while (fetch_nextpage) {

    // Make the Aircall Calls API call and parse the response
    var calls_response = UrlFetchApp.fetch(calls_url, get_params);
    var calls_json = calls_response.getContentText();
    var calls_data = JSON.parse(calls_json);

    // For each call...
    for (var i = 0; i < calls_data.calls.length; i++) {
      // Determine the potential length of the call
      var calltime = ((parseInt((new Date().getTime())/1000) - calls_data.calls[i].answered_at)/60).toFixed(2);

      // Determine the potential waiting time of the call
      var waittime = (parseInt((new Date().getTime())/1000) - calls_data.calls[i].started_at);

      // If the call is still active (user is on a call)...
      if (calls_data.calls[i].status == "answered") {

        // If there exists a user (some very special cases where an answered call could have
        // an empty user object)...
        if (calls_data.calls[i].user) {
          callsdb[calls_data.calls[i].user.id.toString()] = {"direction": calls_data.calls[i].direction,
            "number": calls_data.calls[i].raw_digits, "aircall_number": calls_data.calls[i].number.name,
            "calltime": calltime.toString()};
        }

      // ...else, if the call is in the queue...
      } else if (calls_data.calls[i].status == "initial") {

        // Generate a random color
        var color = Math.floor(Math.random()*16777215).toString(16);

        // Keep track of certain data related to this call in-queue
        callswaitingdb[calls_data.calls[i].number.name] = {
          "aircall_number": calls_data.calls[i].number.name,
          "number": calls_data.calls[i].raw_digits,
          "waittime": waittime.toString(),
          "aircall_number_id": calls_data.calls[i].number.id,
          "aircall_number_users": [], // list of users assigned to this Aircall phonenum
          "color": "#" + color,
          "livefeed_cells": [] // list of spreadsheet cells of users assigned to this Aircall phonenum
        };

        var users_with_numbers = Object.keys(user_numberdb);

        // For each Aircall user...
        users_with_numbers.forEach(function (user) {
          // If we find this Aircall phone number is assigned to this user...
          if (user_numberdb[user][calls_data.calls[i].number.id]) {

            // If there is a team filter...
            if (filtered_teams) {



              if (teamsdb[user]) {

                // Get the list of all Aircall teams assigned to this user
                var user_teams = Object.keys(teamsdb[user]);

                // For each team...
                for (var k = 0; k < user_teams.length; k++) {
                  // If this team is part of the list of filtered teams...
                  if (filtered_teams[user_teams[k]]) {
                    callswaitingdb[calls_data.calls[i].number.name].aircall_number_users.push(user);
                    break;
                  }
                }


              } else if (filtered_teams[NO_TEAM]) {
                callswaitingdb[calls_data.calls[i].number.name].aircall_number_users.push(user);
              }



            // ... else, there is no team filter...
            } else {
              callswaitingdb[calls_data.calls[i].number.name].aircall_number_users.push(user);
            }
          }
        });

        if (filtered_teams && !callswaitingdb[calls_data.calls[i].number.name].
        aircall_number_users.length) {
          delete callswaitingdb[calls_data.calls[i].number.name];
        }

      // ...else, if the call is finished AND is not a missed call...
      } else if (calls_data.calls[i].status == "done" && calls_data.calls[i].user) {

        // Keep track of this finished call for the specific user
        if (!numcallsdb[calls_data.calls[i].user.id.toString()]) {
          numcallsdb[calls_data.calls[i].user.id.toString()] = [0, 0, 0, 0];
        }

        // If this is an inbound call..
        if (calls_data.calls[i].direction == "inbound") {
          numcallsdb[calls_data.calls[i].user.id.toString()][0]++;

          // If there is a defined SLA...
          if (SLA) {
            // Determine the amount of time caller waited before picking up
            var sla_waitingtime = calls_data.calls[i].answered_at - calls_data.calls[i].started_at;

            // Keep track of this time, which will be averaged later on in the code
            numcallsdb[calls_data.calls[i].user.id.toString()][3] += sla_waitingtime;
          }

        // ...else, this is an outbound call...
        } else {
          numcallsdb[calls_data.calls[i].user.id.toString()][1]++;
        }
      // ...else, if the call is a missed call...
      } else if (calls_data.calls[i].missed_call_reason) {

        // Get the list of all Aircall users
        var users_with_numbers = Object.keys(user_numberdb);

        // For each Aircall user...
        users_with_numbers.forEach(function (user) {
          // If this user is assigned to this phone number that missed the call...
          if (user_numberdb[user][calls_data.calls[i].number.id]) {
            // Keep track of this missed call for the specific user
            if (!numcallsdb[user]) {
              numcallsdb[user] = [0, 0, 0, 0];
            }

            numcallsdb[user][2]++;
          }
        });
      }
    }

    // Determine if we need to iterate through other pages of results from the API call...
    if (calls_data.meta.next_page_link) {
      avail_pagenum++;
      calls_url = "https://api.aircall.io/v1/calls?order=desc&per_page=50&page=" + avail_pagenum.toString() +
      "&from=" + mins_before_currtime.toString() + "&to=" + currtime.toString();
      //calls_url = calls_data.meta.next_page_link;
    } else {
      fetch_nextpage = false;
    }
  }

  Logger.log("Queried for all Aircall calls...");



  fetch_nextpage = true;
  avail_pagenum = 1;

  // Define the url for the Aircall User Availabilities API call
  var users_url = "https://api.aircall.io/v1/users/availabilities?order=asc&page=" + avail_pagenum.toString() + "&per_page=50";

  // While there is still at least one page to review for Aircall user information
  // (pagination)...
  while (fetch_nextpage) {

    // Make the Aircall user Availabilities API call and parse the response
    var users_response = UrlFetchApp.fetch(users_url, get_params);
    var users_json = users_response.getContentText();
    var users_data = JSON.parse(users_json);

    // For each user...
    for (var i = 0; i < users_data.users.length; i++) {
      var userid = users_data.users[i].id;
      var userstatus = users_data.users[i].availability;

      // If the user is associated with a team...
      if (teamsdb[userid]) {
        // Keep track of the user (and their availability) and where to place them
        // (under a team) in the Activity Tracker
        var userteams = Object.keys(teamsdb[userid]);
        for (var j = 0; j < userteams.length; j++) {
          // If we created a header in the Activity Tracker for this team (either because
          // there is no team filter or if there is a team filter, this team is
          // part of it)...
          if (livefeed_results[userteams[j]]) {
            livefeed_results[userteams[j]].push(userid + ":" + userstatus);
          }
        }

      // ...else, the user is not associated with a team...
      } else {
        // Keep track of the user (and their availability) and where to place them
        // in the Activity Tracker
        if (livefeed_results[NO_TEAM]) {
          livefeed_results[NO_TEAM].push(userid + ":" + userstatus);
        }
      }
    }

    // Determine if we need to iterate through other pages of results from the API call...
    if (users_data.meta.next_page_link) {
      avail_pagenum++;
      users_url = "https://api.aircall.io/v1/users/availabilities?order=asc&page=" + avail_pagenum.toString() + "&per_page=50";
    } else {
      fetch_nextpage = false;
    }
  }

  Logger.log("Queried for all Aircall user availabilities...");


  // Clear the content and format of the Activity Tracker
  livefeed_sheet.clear();

  // Counters to keep track of where we are within the "Activity Tracker" sheet as we start displaying
  // user availability information
  var ss_def_rownum = 1;
  var ss_pos_rownum = 1;
  var ss_pos_colnum = 1;

  callswaiting_nums = Object.keys(callswaitingdb);

  // If there are calls in queue...
  if (!(callswaiting_nums.length == 0)) {

    // Format and populate the header in the Activity Tracker for calls in queue
    var wait_cell = formatWaitCell(ss_pos_rownum, ss_pos_colnum, livefeed_sheet,
      WAIT_CELL_ROWSIZE, WAIT_CELL_COLSIZE, "header", {});
    ss_pos_colnum += wait_cell.getNumColumns();

    // For each call waiting in queue...
    callswaiting_nums.forEach(function (waitnum) {
      // Format and populate a cell within the Activity Tracker
      wait_cell = formatWaitCell(ss_pos_rownum, ss_pos_colnum, livefeed_sheet,
        WAIT_CELL_ROWSIZE, WAIT_CELL_COLSIZE, "", callswaitingdb[waitnum]);

      ss_pos_colnum += wait_cell.getNumColumns();
    });

    // After all calls waiting in queue have been displayed in the Activity Tracker, reset our
    // position in the spreadsheet -- we will start displaying user availabilities next
    ss_pos_rownum += (WAIT_CELL_ROWSIZE + DIVIDER_ROWSIZE);
    ss_def_rownum = ss_pos_rownum;
    ss_pos_colnum = 1;
  }

  Logger.log("Finished displaying all calls in queue...");


  // To optimize performance, we will keep track of each cell that represents a Team Name
  // so we can batch format them
  var teams_format_list = [];

  // To optimize performance, we will keep track of each cell that represents a User and their
  // specific availability status
  // Structure: [
  //    [List of User Cells - Available],
  //    [List of User Cells - Do Not Disturb],
  //    [List of User Cells - Offline],
  //    [List of User Cells - After Call Work],
  //    [List of User Cells - In Call]
  // ]
  var users_format_list = [[], [], [], [], []];

  // To optimize performance, we will keep track of each cell that represents a summary of
  // user call history over the interval period
  var nums_format_list = [];

  // To optimize performance, we will keep track of each cell that represents a summary of
  // an active call (user is still on the call)
  var incall_format_list = [];

  // Keep track of all users who are above their SLA times
  var sla_format_list = [];

  livefeed_results_teams = Object.keys(livefeed_results);

  // For each Team Name...
  livefeed_results_teams.forEach(function (teamid) {
    var teamname = teamnamesdb[teamid.toString()];
    var teamusers = livefeed_results[teamid];

    // Update the cell by setting the Team Name header
    var team_cell = livefeed_sheet.getRange(ss_pos_rownum, ss_pos_colnum,
      TEAM_CELL_ROWSIZE, TEAM_CELL_COLSIZE).merge().setValue(teamname);
    teams_format_list.push(team_cell.getA1Notation());

    ss_pos_rownum += team_cell.getNumRows();

    // For each user that belongs to this team...
    teamusers.forEach(function (team_user) {
      var userid = team_user.split(":")[0];
      var username = userdb[userid];

      // If the user ID cannot be mapped...
      if (!username) {
        username = "NEW AGENT (" + userid + ")";
      }

      // Get the user's availability status
      var userstatus = team_user.split(":")[1];

      // Update the cell by setting the username
      var user_cell = livefeed_sheet.getRange(ss_pos_rownum, ss_pos_colnum,
        USER_CELL_ROWSIZE, USER_CELL_COLSIZE).merge().setValue(username);


      callswaiting_nums = Object.keys(callswaitingdb);

      // If there are calls in-queue...
      if (!(callswaiting_nums.length == 0)) {
        // For each call waiting in-queue...
        callswaiting_nums.forEach(function (waitnum) {
          // Retrieve the list of Aircall users assigned to the Aircall phonenum
          // which this call is waiting on
          var calls_waiting_num_users = callswaitingdb[waitnum].aircall_number_users;

          var found = calls_waiting_num_users.find(function(id) {
            return id == userid;
          });

          // If we find this user is assigned to this Aircall phonenum, keep track of it
          // for formatting purposes...
          if (found) {
            callswaitingdb[waitnum].livefeed_cells.push(user_cell.getA1Notation());
          }
        });
      }

      ss_pos_rownum += user_cell.getNumRows();

      // If the user was part of any inbound/outbound/missed calls within the interval period, we
      // will display it in a cell...
      if (numcallsdb[userid]) {
        var userval = numcallsdb[userid][0] + " INB / " + numcallsdb[userid][1] + " OUT";

        if (user_num_mappings_enabled) {
          userval += " / " + numcallsdb[userid][2] + " MISS";
        }

        // If there is a defined SLA...
        if (SLA && numcallsdb[userid][3]) {
          // Determine the average time for this user to pick up a call
          var avg_sla_waitingtime = parseInt(numcallsdb[userid][3]/numcallsdb[userid][0]);

          // If the average time is greater than the defined SLA, keep track of the user for formatting
          // within the Activity Tracker
          if (avg_sla_waitingtime > SLA) {
            sla_format_list.push(user_cell.getA1Notation());
          }
        }

        var user_nums_cell = livefeed_sheet.getRange(ss_pos_rownum, ss_pos_colnum,
          NUM_SUMMARY_ROWSIZE, NUM_SUMMARY_COLSIZE).merge().setValue(userval);
        nums_format_list.push(user_nums_cell.getA1Notation());

        ss_pos_rownum += user_nums_cell.getNumRows();
      }

      // Depending on the user's status, we will add them to the appropriate list for batch
      // formatting later...
      switch (userstatus) {
        case "available":
          users_format_list[0].push(user_cell.getA1Notation());
          break;
        case "do_not_disturb":
          users_format_list[1].push(user_cell.getA1Notation());
          break;
        case "offline":
          users_format_list[2].push(user_cell.getA1Notation());
          break;
        case "after_call_work":
          users_format_list[3].push(user_cell.getA1Notation());
          break;
        case "in_call":
          users_format_list[4].push(user_cell.getA1Notation());

          // If a user is in a call, we will display it in a cell...
          if (callsdb[userid]) {
            user_cell = formatInCallCell(ss_pos_rownum, ss_pos_colnum, livefeed_sheet, callsdb[userid]);
            incall_format_list.push(user_cell.getA1Notation());

            ss_pos_rownum += user_cell.getNumRows();
          }
          break;
      }

    });

    // Re-position ourselves within the "Activity Tracker" sheet to start displaying the next team
    ss_pos_rownum = ss_def_rownum;
    ss_pos_colnum += team_cell.getNumColumns();
  });

  Logger.log("Finished organizing rest of results to be displayed...");



  // If there are Team Name headers that require batch formatting...
  if (!(teams_format_list.length == 0)) {
    livefeed_sheet.getRangeList(teams_format_list).setHorizontalAlignment("center");
    livefeed_sheet.getRangeList(teams_format_list).setVerticalAlignment("middle");
    livefeed_sheet.getRangeList(teams_format_list).setFontWeight("bold");
    livefeed_sheet.getRangeList(teams_format_list).setBackground("#006b51");
    livefeed_sheet.getRangeList(teams_format_list).setFontColor("white");
    livefeed_sheet.getRangeList(teams_format_list).setBorder(true, true, true, true, false, false,
    "white", SpreadsheetApp.BorderStyle.SOLID);

  }

  // If there are user cells that are "available" status and require batch formatting...
  if (!(users_format_list[0].length == 0)) {
    livefeed_sheet.getRangeList(users_format_list[0]).setHorizontalAlignment("center");
    livefeed_sheet.getRangeList(users_format_list[0]).setVerticalAlignment("middle");
    livefeed_sheet.getRangeList(users_format_list[0]).setBackground("#d8f3ec");
    livefeed_sheet.getRangeList(users_format_list[0]).setFontColor("#00b388");
    livefeed_sheet.getRangeList(users_format_list[0]).setBorder(true, true, true, true, false, false,
    "white", SpreadsheetApp.BorderStyle.SOLID);
  }

  // If there are user cells that are "do_not_disturb" status and require batch formatting...
  if (!(users_format_list[1].length == 0)) {
    livefeed_sheet.getRangeList(users_format_list[1]).setHorizontalAlignment("center");
    livefeed_sheet.getRangeList(users_format_list[1]).setVerticalAlignment("middle");
    livefeed_sheet.getRangeList(users_format_list[1]).setBackground("#ff7b7b");
    livefeed_sheet.getRangeList(users_format_list[1]).setFontColor("white");
    livefeed_sheet.getRangeList(users_format_list[1]).setBorder(true, true, true, true, false, false,
    "white", SpreadsheetApp.BorderStyle.SOLID);
  }

  // If there are user cells that are "offline" status and require batch formatting...
  if (!(users_format_list[2].length == 0)) {
    livefeed_sheet.getRangeList(users_format_list[2]).setHorizontalAlignment("center");
    livefeed_sheet.getRangeList(users_format_list[2]).setVerticalAlignment("middle");
    livefeed_sheet.getRangeList(users_format_list[2]).setBackground("#eeeeee");
    livefeed_sheet.getRangeList(users_format_list[2]).setFontColor("#595959");
    livefeed_sheet.getRangeList(users_format_list[2]).setBorder(true, true, true, true, false, false,
    "white", SpreadsheetApp.BorderStyle.SOLID);
  }

  // If there are user cells that are "after_call_work" status and require batch formatting...
  if (!(users_format_list[3].length == 0)) {
    livefeed_sheet.getRangeList(users_format_list[3]).setHorizontalAlignment("center");
    livefeed_sheet.getRangeList(users_format_list[3]).setVerticalAlignment("middle");
    livefeed_sheet.getRangeList(users_format_list[3]).setBackground("#ffe76d");
    livefeed_sheet.getRangeList(users_format_list[3]).setFontColor("orange");
    livefeed_sheet.getRangeList(users_format_list[3]).setBorder(true, true, true, true, false, false,
    "white", SpreadsheetApp.BorderStyle.SOLID);
  }

  // If there are user cells that are "in_line" status and require batch formatting...
  if (!(users_format_list[4].length == 0)) {
    livefeed_sheet.getRangeList(users_format_list[4]).setHorizontalAlignment("center");
    livefeed_sheet.getRangeList(users_format_list[4]).setVerticalAlignment("middle");
    livefeed_sheet.getRangeList(users_format_list[4]).setBackground("orange");
    livefeed_sheet.getRangeList(users_format_list[4]).setFontColor("white");
    livefeed_sheet.getRangeList(users_format_list[4]).setBorder(true, true, true, true, false, false,
    "white", SpreadsheetApp.BorderStyle.SOLID);
  }

  // If there are cells that represent inbound/outbound call summaries that require batch formatting...
  if (!(nums_format_list.length == 0)) {
    livefeed_sheet.getRangeList(nums_format_list).setHorizontalAlignment("center");
    livefeed_sheet.getRangeList(nums_format_list).setVerticalAlignment("middle");
    livefeed_sheet.getRangeList(nums_format_list).setBorder(true, true, true, true, false, false,
    "white", SpreadsheetApp.BorderStyle.SOLID);
    livefeed_sheet.getRangeList(nums_format_list).setFontStyle("italic");
    livefeed_sheet.getRangeList(nums_format_list).setBackground("#b7b7b7");
    livefeed_sheet.getRangeList(nums_format_list).setFontColor("white");
  }

  // If there are cells that represent active call summaries that require batch formatting...
  if (!(incall_format_list.length == 0)) {
    livefeed_sheet.getRangeList(incall_format_list).setHorizontalAlignment("center");
    livefeed_sheet.getRangeList(incall_format_list).setVerticalAlignment("middle");
    livefeed_sheet.getRangeList(incall_format_list).setBorder(true, true, true, true, false, false,
    "orange", SpreadsheetApp.BorderStyle.DASHED);
    livefeed_sheet.getRangeList(incall_format_list).setBackground("#ffeba6");
    livefeed_sheet.getRangeList(incall_format_list).setFontColor("orange");
  }

  // If there are cells that represent Aircall users who are not meeting defined SLA...
  if (!(sla_format_list.length == 0)) {
    livefeed_sheet.getRangeList(sla_format_list).setFontWeight("bold");
    livefeed_sheet.getRangeList(sla_format_list).setFontColor("red");
  }

  callswaiting_nums = Object.keys(callswaitingdb);

  // If there are calls in queue...
  if (!(callswaiting_nums.length == 0)) {
    // For each call waiting in-queue...
    callswaiting_nums.forEach(function (waitnum) {
      // If we find at least one user assigned to the Aircall phonenum which this
      // call is waiting on, format it...
      if (callswaitingdb[waitnum].livefeed_cells.length > 0) {
        livefeed_sheet.getRangeList(callswaitingdb[waitnum].livefeed_cells).
          setBorder(true, true, true, true, false, false, callswaitingdb[waitnum].color,
          SpreadsheetApp.BorderStyle.SOLID);
      }
    });
  }


  // Write all the changes to the spreadsheet
  SpreadsheetApp.flush();

  Logger.log("Displayed all results...");
}


/*
* Function that formats a cell with a username inside within the Activity Tracker
*   @param  {Integer} row_num     Row number within the sheet
*           {Integer} col_num     Column number within the sheet
*           {Sheet} sheet         The Sheet object where we are going to format the cell
*   @return {Range} cell          The Range object which represents the cell we formatted
*/
function formatUserCell(row_num, col_num, sheet) {
  var cell;
  cell = sheet.getRange(row_num, col_num, 2, 2).merge();
  return cell;
}


/*
* Function that formats a cell that has a summary of information about an
* active call within the Activity Tracker
*   @param  {Integer} row_num     Row number within the sheet
*           {Integer} col_num     Column number within the sheet
*           {Sheet} sheet         The Sheet object where we are going to format the cell
*           {JSON} info_params    Data structure which contains a summary of information
*                                 about the active call
*   @return {Range} cell          The Range object which represents the cell we formatted
*/
function formatInCallCell (row_num, col_num, sheet, info_params) {
  const INCALL_CELL_ROWSIZE = 3;
  const INCALL_CELL_COLSIZE = 2;

  var cell = sheet.getRange(row_num, col_num, INCALL_CELL_ROWSIZE, INCALL_CELL_COLSIZE).merge();

  var direction = "";
  if (info_params["direction"] == "outbound") {
    direction = "OUT";
  } else {
    direction = "INB";
  }

  cell.setValue(info_params["number"] + "   (" + direction + ")" + "\n" + info_params["aircall_number"] +
    "\n" + info_params["calltime"] + " mins");

  return cell;
}


/*
* Function that formats a cell that represents a call waiting in queue
* within the Activity Tracker
*   @param  {Integer} row_num     Row number within the sheet
*           {Integer} col_num     Column number within the sheet
*           {Sheet} sheet         The Sheet object where we are going to format the cell
*           {Integer} rowsize     Number of rows that should be merged for the cell
*           {Integer} colsize     Number of columns that should be merged for the cell
*           {String} type         Either the header or a standard cell
*           {JSON} info_params    Data structure which contains a summary of information
*                                 about the queued call
*   @return {Range} cell          The Range object which represents the cell we formatted
*/
function formatWaitCell(row_num, col_num, sheet, rowsize, colsize, type, info_params) {
  var cell = sheet.getRange(row_num, col_num, rowsize, colsize).merge();
  cell.setHorizontalAlignment("center");
  cell.setVerticalAlignment("middle");

  var cache = PropertiesService.getScriptProperties();

  // If we are formatting the header of the section representing all calls waiting in queue...
  if (type == "header") {
    cell.setBorder(false, true, true, true, false, false, "white", SpreadsheetApp.BorderStyle.SOLID);
    cell.setBackground("#00b388");
    cell.setFontColor("white");
    cell.setFontWeight("bold");
    cell.setValue("Live Queue");

  // ...else, this is a standard cell which will contain information about a queued call...
  } else {
    var waitlevel1 = parseInt(cache.getProperty("wait_lvl_1"));
    var waitlevel2 = parseInt(cache.getProperty("wait_lvl_2"));

    // Determine whether the waiting time of the queued call is within the accepted thresholds,
    // and format the cell according to the length of waiting time...
    if (parseInt(info_params["waittime"]) < waitlevel1) {
      cell.setBackground("#d8f3ec");
      cell.setFontColor("#00b388");
      cell.setBorder(false, true, true, true, false, false, "#34a853", SpreadsheetApp.BorderStyle.DASHED);
    } else if (parseInt(info_params["waittime"]) < waitlevel2) {
      cell.setBackground("#ffeba6");
      cell.setFontColor("orange");
      cell.setBorder(false, true, true, true, false, false, "orange", SpreadsheetApp.BorderStyle.DASHED);
    } else {
      cell.setBackground("#ff7b7b");
      cell.setFontColor("white");
      cell.setBorder(false, true, true, true, false, false, "white", SpreadsheetApp.BorderStyle.DASHED);
    }

    // If the feature to highlight which users could potentially pick up this
    // call in-queue is enabled...
    if (info_params["aircall_number_users"].length > 0) {
      cell.setBorder(false, true, true, true, false, false, info_params["color"],
        SpreadsheetApp.BorderStyle.SOLID);
    }

    cell.setValue(info_params["number"] + "\n" + info_params["aircall_number"] +
      "\n Waiting for " + info_params["waittime"] + "s...");
  }

  return cell;
}


/*
* Setup the entire spreadsheet for the Activity Tracker by estimating how many columns and rows will
* need to be created to store Aircall user availabilities.
*/
function setupActivityTracker() {

  // Set various constants that will help us estimate how many columns/rows
  // we need in various sheets
  const WIDTH_BUFFER = 1.25;
  const HEIGHT_BUFFER = 1.15;
  const AVG_ON_CALL = 0.5;
  const MAX_PERC_OF_USERS_ON_TEAM = 0.5;
  const TEAM_CELL_ROWSIZE = 4;
  const TEAM_CELL_COLSIZE = 2;
  const USER_CELL_ROWSIZE = 2;
  const INCALL_CELL_ROWSIZE = 3;
  const NUM_SUMMARY_ROWSIZE = 1;
  const DEF_NUM_COLS = 15;
  const DEF_NUM_ROWS = 150

  var cache = PropertiesService.getScriptProperties();
  var livefeed_ss = SpreadsheetApp.getActiveSpreadsheet();

  Logger.log("Creating 'Activity Tracker' sheet...");

  // Create the "Activity Tracker" sheet if required
  var livefeed_sheet;
  if (!livefeed_ss.getSheetByName("Activity Tracker")) {
    livefeed_sheet = livefeed_ss.insertSheet().setName("Activity Tracker");
  } else {
    livefeed_sheet = livefeed_ss.getSheetByName("Activity Tracker");
  }

  // Clear the "Activity Tracker" sheet of any content and format
  livefeed_sheet.clear();

  // Get the API Key from the cache
  var API_KEY;
  if (cache.getProperty("api_key")) {API_KEY = cache.getProperty("api_key");}

  // If no API key is still set, exit the program
  if (!API_KEY) {
    SpreadsheetApp.getUi().alert("No API key set, will not continue.");
    return;
  }

  // Set headers and parameters for any future Aircall API calls
  var headers = {
    "Authorization" : "Basic " + API_KEY
  };
  var get_params = {
    "method":"GET",
    "headers":headers
  };

  // Define the url for the Aircall User API call
  var users_url = "https://api.aircall.io/v1/users?per_page=50";
  // Define the url for the Aircall Team API call
  var teams_url = "https://api.aircall.io/v1/teams?per_page=50";
  // Define the url for the Aircall Number API call
  var numbers_url = "https://api.aircall.io/v1/numbers?per_page=50";

  // Make the Aircall User API call and parse the response
  var users_response = UrlFetchApp.fetch(users_url, get_params);
  var users_json = users_response.getContentText();
  var users_data = JSON.parse(users_json);

  // Make the Aircall Team API call and parse the response
  var teams_response = UrlFetchApp.fetch(teams_url, get_params);
  var teams_json = teams_response.getContentText();
  var teams_data = JSON.parse(teams_json);

  // Make the Aircall Number API call and parse the response
  var numbers_response = UrlFetchApp.fetch(numbers_url, get_params);
  var numbers_json = numbers_response.getContentText();
  var numbers_data = JSON.parse(numbers_json);

  // Retrieve the total number of Aircall users, teams, and numbers in this instance
  var total_users = users_data.meta.total;
  var total_teams = teams_data.meta.total + 1;
  var total_numbers = numbers_data.meta.total;

  // Delete all columns and rows greater than the default number of rows
  // and columns in the sheet
  var max_col = livefeed_sheet.getMaxColumns();
  var max_row = livefeed_sheet.getMaxRows();
  if ((max_col - DEF_NUM_ROWS) > 0) {
    livefeed_sheet.deleteColumns(DEF_NUM_ROWS, (max_col - DEF_NUM_ROWS));
  }
  if ((max_row - DEF_NUM_COLS) > 0) {
    livefeed_sheet.deleteRows(DEF_NUM_COLS, (max_row - DEF_NUM_COLS));
  }

  // Find the last column and row index
  max_col = livefeed_sheet.getMaxColumns();
  max_row = livefeed_sheet.getMaxRows();

  // Estimate the number of columns required to add to the "Activity Tracker" sheet
  var num_columns = Math.ceil((total_teams * TEAM_CELL_COLSIZE) * WIDTH_BUFFER) - max_col;
  if (num_columns <= 0) {num_columns = 1;}

  // Estimate the number of rows required to add to the "Activity Tracker" sheet
  var num_rows = Math.ceil(((TEAM_CELL_ROWSIZE) + ((total_users * MAX_PERC_OF_USERS_ON_TEAM) * USER_CELL_ROWSIZE) +
    ((total_users * AVG_ON_CALL) * (INCALL_CELL_ROWSIZE + NUM_SUMMARY_ROWSIZE))) *
    HEIGHT_BUFFER) - max_row;
  if (num_rows <= 0) {num_rows = 1;}

  // Create the columns and rows required
  livefeed_sheet.insertColumns(1, num_columns);
  livefeed_sheet.insertRows(1, num_rows);


  Logger.log("Creating 'UserDB' sheet...");

  // Create the "UserDB" sheet if required
  var userdb_sheet;
  if (!livefeed_ss.getSheetByName("UserDB")) {
    userdb_sheet = livefeed_ss.insertSheet().setName("UserDB");
  } else {
    userdb_sheet = livefeed_ss.getSheetByName("UserDB");
  }

  // Clear the "UserDB" sheet of any content or format
  userdb_sheet.clear();


  // Delete all columns and rows greater than the default number of rows
  // and columns in the sheet
  max_col = userdb_sheet.getMaxColumns();
  max_row = userdb_sheet.getMaxRows();

  if ((max_col - DEF_NUM_ROWS) > 0) {
    userdb_sheet.deleteColumns(DEF_NUM_ROWS, (max_col - DEF_NUM_ROWS));
  }
  if ((max_row - DEF_NUM_COLS) > 0) {
    userdb_sheet.deleteRows(DEF_NUM_COLS, (max_row - DEF_NUM_COLS));
  }

  // Find the last column and row index
  max_col = userdb_sheet.getMaxColumns();
  max_row = userdb_sheet.getMaxRows();

  // Estimate the number of rows required to add to sheet in order to store
  // all usernames and user IDs
  var num_rows = Math.ceil(total_users * HEIGHT_BUFFER) - max_row;
  if (num_rows <= 0) {num_rows = 1;}

  // Create the rows required
  userdb_sheet.insertRows(1, num_rows);


  Logger.log("Creating 'NumberDB' sheet...");

  // Create the "NumberDB" sheet if required
  var numberdb_sheet;
  if (!livefeed_ss.getSheetByName("NumberDB")) {
    numberdb_sheet = livefeed_ss.insertSheet().setName("NumberDB");
  } else {
    numberdb_sheet = livefeed_ss.getSheetByName("NumberDB");
  }

  // Clear the "NumberDB" sheet of any content or format
  numberdb_sheet.clear();


  // Delete all columns and rows greater than the default number of rows
  // and columns in the sheet
  max_col = numberdb_sheet.getMaxColumns();
  max_row = numberdb_sheet.getMaxRows();

  if ((max_col - DEF_NUM_ROWS) > 0) {
    numberdb_sheet.deleteColumns(DEF_NUM_ROWS, (max_col - DEF_NUM_ROWS));
  }
  if ((max_row - DEF_NUM_COLS) > 0) {
    numberdb_sheet.deleteRows(DEF_NUM_COLS, (max_row - DEF_NUM_COLS));
  }

  // Find the last column and row index
  max_col = numberdb_sheet.getMaxColumns();
  max_row = numberdb_sheet.getMaxRows();

  // Estimate the number of rows required to add to sheet in order to store
  // all number names and IDs
  var num_rows = Math.ceil(total_numbers * HEIGHT_BUFFER) - max_row;
  if (num_rows <= 0) {num_rows = 1;}

  // Create the rows required
  numberdb_sheet.insertRows(1, num_rows);

  Logger.log("Creating 'UserSelection' sheet...");

  // Create the "UserSelection" sheet if required
  var userselection_sheet;
  if (!livefeed_ss.getSheetByName("UserSelection")) {
    userselection_sheet = livefeed_ss.insertSheet().setName("UserSelection");
  } else {
    userselection_sheet = livefeed_ss.getSheetByName("UserSelection");
  }

  // Clear the "UserSelection" sheet of any content or format
  userselection_sheet.clear();

  // Ensures that all Aircall users and teams are refreshed after setup
  if (cache.getProperty("team_cache_time")) {
    cache.deleteProperty("team_cache_time");
  }
  if (cache.getProperty("user_cache_time")) {
    cache.deleteProperty("user_cache_time");
  }
}


/*
* Function that gets called as soon as the spreadsheet is opened
*/
function onOpen() {
  // Create a custom menu button to show call history of a select user or all users
  SpreadsheetApp
   .getUi()
   .createMenu("Calls")
   .addItem("Select User", "showUserCallsSidebar")
   .addItem("All Users", "showAllCallsSidebar")
   .addItem("Advanced Search", "showAdvancedSearchSidebar")
   .addToUi();

  // Create a custom menu button to show actions related to the Activity Tracker
  SpreadsheetApp
   .getUi()
   .createMenu("Activity Tracker")
   .addItem("Refresh", "activityTracker")
   .addItem("Configure", "showConfigSidebar")
   .addItem("Setup Activity Tracker", "setupActivityTracker")
   .addItem("Map Users to Numbers", "mapNumbersToUsers")
   .addItem("Clear Cache", "clearCache")
   .addItem("Switch User Status", "switchUserStatus")
   .addToUi();
}


/*
* Function that creates a sidebar containing call history for a select user
*/
function showUserCallsSidebar() {

  // Gather all the user information that maps User ID to Username
  var userdb_reverse = {};
  userdb_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UserDB");
  var userdb_reverse_data = userdb_sheet.getDataRange().getValues();
  for (var i = 0; i < userdb_reverse_data.length; i++) {
    userdb_reverse[userdb_reverse_data[i][1].toString()] = userdb_reverse_data[i][0];
  }

  // Get the selected user's ID
  var username = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UserSelection").getRange("A1").getValue();
  var userid = userdb_reverse[username];

  // If we find the User ID...
  if (userid) {
    // Populate a separate sheet that will contain the User ID -- this will be extracted later when the
    // sidebar HTML page is generated
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UserSelection").getRange("B1").setValue(userid);

    // Create an HTML page from the "UserCallHistory" template and display it in the sidebar
    var widget = HtmlService.createTemplateFromFile('UserCallHistory').evaluate();
    widget.setTitle("Review Call History - " + username);
    SpreadsheetApp.getUi().showSidebar(widget);
  } else {
    SpreadsheetApp.getUi().alert("Cannot determine User ID.");
  }
}


/*
* Function that creates a sidebar containing call history for all users
*/
function showAllCallsSidebar() {
  // Create an HTML page from the "AllCallHistory" template and display it in the sidebar
  var widget = HtmlService.createTemplateFromFile('AllCallHistory').evaluate();
  widget.setTitle("Review Call History");
  SpreadsheetApp.getUi().showSidebar(widget);
}


/*
*
*/
function showAdvancedSearchSidebar() {
  // Create an HTML page from the "AllCallHistory" template and display it in the sidebar
  var widget = HtmlService.createTemplateFromFile('AdvancedSearch').evaluate();
  widget.setTitle("Advanced Search");
  SpreadsheetApp.getUi().showSidebar(widget);
}



/*
* Function that creates a sidebar containing Activity Tracker configuration input
*/
function showConfigSidebar() {

  var widget = HtmlService.createTemplateFromFile('SetConfig').evaluate();
  widget.setTitle("Activity Tracker Config Parameters");
  SpreadsheetApp.getUi().showSidebar(widget);
}


/*
* Function that gets called when a cell is selected on-screen
*/
function onSelectionChange(e) {
  var active_sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();

  // Ensure we are only selecting values from the "Activity Tracker" sheet...
  if (active_sheet == "Activity Tracker") {
    // Get the range information of the current cell being selected
    var range = e.range;
    var val = range.getValue();

    // Populate a separate spreadsheet which will keep track of the user cell that is currently
    // selected
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UserSelection").getRange("A1").setValue(val);
  }
}


/*
* Function that takes the input configuration params from the config param sidebar
* and loads them into the cache
*/
function setConfigParams(interval, api_key, user_interval, team_interval,
wait_lvl_1, wait_lvl_2, sla, filtered_teams) {
  Logger.log("Setting new activity period to: " + parseInt(interval).toFixed(0));
  Logger.log("Setting new API Key to: " + api_key);
  Logger.log("Setting new user refresh interval to: " + parseInt(user_interval).toFixed(0));
  Logger.log("Setting new team refresh interval to: " + parseInt(team_interval).toFixed(0));
  Logger.log("Setting new wait level 1 threshold to: " + parseInt(wait_lvl_1).toFixed(0));
  Logger.log("Setting new wait level 2 threshold to: " + parseInt(wait_lvl_2).toFixed(0));
  Logger.log("Setting new wait SLA threshold to: " + parseInt(sla).toFixed(0));
  Logger.log("Setting new teams filtasdasdasdasder to: " + filtered_teams);

  var cache = PropertiesService.getScriptProperties();

  // If a team filter was set, retrieve the list of teams...
  if (filtered_teams) {
    var teams = filtered_teams.split(";");
  } else {
    var teams = [];
  }

  // Create a JSON object which will contain every team included in the
  // team filter as a key
  var teams_json = {};
  for (var i = 0; i < teams.length; i++) {
    Logger.log("Entering team: " + teams[i]);
    teams_json[teams[i]] = true;
  }

  if (interval) {cache.setProperty("interval", parseInt(interval).toFixed(0));}

  // If there is an API key set...
  if (api_key) {
    // If the API key set is different than the previous, we will need
    // to force a refresh of all data within the Activity Tracker
    if (cache.getProperty("api_key") != api_key) {
      cache.setProperty("api_key_changed", "1");
    } else {
      cache.setProperty("api_key_changed", "0");
    }
    cache.setProperty("api_key", api_key);
  }

  if (user_interval) {cache.setProperty("user_interval", parseInt(user_interval).toFixed(0));}
  if (team_interval) {cache.setProperty("team_interval", parseInt(team_interval).toFixed(0));}
  if (wait_lvl_1) {cache.setProperty("wait_lvl_1", parseInt(wait_lvl_1).toFixed(0));}
  if (wait_lvl_2) {cache.setProperty("wait_lvl_2", parseInt(wait_lvl_2).toFixed(0));}
  if (sla) {cache.setProperty("sla", parseInt(sla).toFixed(0));}

  // If the list of filtered teams is different than the team filter set previously, keep
  // track so we can force-refresh the team info
  if (cache.getProperty("filteredteams_cache_db") != JSON.stringify(teams_json)) {
    cache.setProperty("filteredteams_changed", "1");
  } else {
    cache.setProperty("filteredteams_changed", "0");
  }

  // If there is a team filter...
  if (Object.keys(teams_json).length) {
    cache.setProperty("filteredteams_cache_db", JSON.stringify(teams_json));

  // ...else, there is no team filter set...
  } else {
    cache.deleteProperty("filteredteams_cache_db");
    Logger.log("Deleting team filter...");
  }
}


/*
* Function that switches the selected user's availability status
*/
function switchUserStatus() {

  // Gather all the user information that maps User ID to Username
  var userdb_reverse = {};
  var userdb_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UserDB");
  var userdb_reverse_data = userdb_sheet.getDataRange().getValues();
  for (var i = 0; i < userdb_reverse_data.length; i++) {
    userdb_reverse[userdb_reverse_data[i][1].toString()] = userdb_reverse_data[i][0];
  }

  // Get the selected user's ID
  var username = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UserSelection").getRange("A1").getValue();
  var userid = userdb_reverse[username];

  var cache = PropertiesService.getScriptProperties();

  // If we find the User ID...
  if (userid) {
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UserSelection").getRange("B1").setValue(userid);

    // Set headers and parameters for any future Aircall API calls
    var headers = {
      "Authorization" : "Basic " + cache.getProperty("api_key")
    };

    var get_params = {
      "method":"GET",
      "headers":headers
    };

    var put_params = {
      "method":"PUT",
      "contentType": "application/json",
      "headers": headers,
      "payload": ""
    };

    var user_avail_url = "https://api.aircall.io/v1/users/" + userid;

    // Fetch the user's current availability status
    var user_avail_response = UrlFetchApp.fetch(user_avail_url, get_params);
    var user_avail_json = user_avail_response.getContentText();
    var user_avail_data = JSON.parse(user_avail_json);

    // If the customer does not have availability status set to "custom"...
    if (user_avail_data.user.availability_status != "custom") {
      // Determine the new availability status
      var new_avail_status;
      if (user_avail_data.user.availability_status == "available") {
        new_avail_status = "unavailable";
      } else {
        new_avail_status = "available";
      }

      Logger.log("Setting user availability status from '" +
        user_avail_data.user.availability_status + "' to '" + new_avail_status + "'...");

      // Update the user's current availability status to the new one
      put_params.payload = '{"availability_status": "' + new_avail_status + '"}';
      user_avail_response = UrlFetchApp.fetch(user_avail_url, put_params);

      // If the update was not successful...
      if (user_avail_response.getResponseCode() !== 200) {
        SpreadsheetApp.getUi().alert("Not able to successfully change user status!");

      // ..., else the update was successful...
      } else {
        // Prompt for refresh of Activity Tracker since the user's availability status changed
        var ui = SpreadsheetApp.getUi();
        var result = ui.alert(
          'User Status Change',
          'Do you want to refresh the Activity Tracker to see updated user availabilities?',
          ui.ButtonSet.YES_NO);

        if (result == ui.Button.YES) {
          activityTracker();
        }
      }
    } else {
      SpreadsheetApp.getUi().alert("User has 'Custom' status set, will not change availability.");
    }
  } else {
    SpreadsheetApp.getUi().alert("Cannot determine User ID.");
  }
}


/*
* Function that clears the cache (resets user refresh period, team refresh period, previously used
* API key, and any other params set through the config param sidebar)
*/
function clearCache() {
  var cache = PropertiesService.getScriptProperties();
  cache.deleteAllProperties();
}


/*
* Server-side function that gets called by the "Advances Search" webpage to make an Aircall API
* call to retrieve call activities within an activity period.
*   @param  {Integer} start_date            The start date (in epoch UNIX time) of activity period
*           {Integer} end_date              The end date (in epoch UNIX time) of activity period
*   @return {String / String[]} all_calls   The result of the API call
*/
function getAdvancedSearchCalls(start_date, end_date) {
  // Get the API key
  var cache = PropertiesService.getScriptProperties();
  const api_key = cache.getProperty("api_key");

  // Set headers and parameters for any future Aircall API calls
  var headers = {
    "Authorization" : "Basic " + api_key
  };
  var get_params = {
    "method":"GET",
    "headers":headers
  };

  var fetch_nextpage = true;

  // Define the url for the Aircall Calls API call
  var calls_url = "https://api.aircall.io/v1/calls?order=desc&per_page=50&from=" +
    start_date.toString() + "&to=" + end_date.toString();

  var avail_pagenum = 1;
  var all_calls = [];

  // If no API key has been set...
  if (!api_key) {
    fetch_nextpage = false;
    all_calls = "No API key has been set!";
  }

  // While there is still at least one page to review for Aircall call information
  // (pagination)...
  while (fetch_nextpage) {
    // Make the Aircall Calls API call and parse the response
    var calls_response = UrlFetchApp.fetch(calls_url, get_params);

    // If we received a safe response from the API call...
    if (calls_response.getResponseCode() == 200) {
      var calls_json = calls_response.getContentText();
      var calls_data = JSON.parse(calls_json);

      all_calls.push(calls_json);

      // Determine if we need to iterate through other pages of results from the API call...
      if (calls_data.meta.next_page_link) {
        avail_pagenum++;
        calls_url = "https://api.aircall.io/v1/calls?order=desc&per_page=50&page=" + avail_pagenum.toString() +
          "&from=" + start_date.toString() + "&to=" + end_date.toString();
        //calls_url = calls_data.meta.next_page_link;
      } else {
        fetch_nextpage = false;
      }

    // ..., else we received a bad response from the API call...
    } else {
      fetch_nextpage = false;
      all_calls = "Bad response from API call!";
    }
  }

  return all_calls;
}


/*
* Function that maps users to phone numbers and places this information within
* the "UserDB" sheet. This setup is useful for mapping calls waiting in-queue
* to specific users within the Activity Tracker (and will also be useful for new
* features that rely on this mapping).
*/
function mapNumbersToUsers() {
  // Get the API key
  var cache = PropertiesService.getScriptProperties();
  const api_key = cache.getProperty("api_key");

  // API calls per minute is 60
  const MAX_CALLS = 55;

  Logger.log("Mapping Aircall users to phone numbers...");

  // Set headers and parameters for any future Aircall API calls
  var headers = {
  "Authorization" : "Basic " + api_key
  };
  var get_params = {
  "method":"GET",
  "headers":headers
  };

  // Gather all the user information that maps User ID to Username
  userdb_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UserDB");
  userdb_sheet.getRange("C:C").clearContent();

  var userdb_data = userdb_sheet.getRange("A:B").getValues();

  //var userdb_data = userdb_sheet.getDataRange().getValues();

  // Create a an array which will keep track of all phone numbers assigned to
  // oeach Aircall user. The list of phone numbers for each user will look like
  // "Phone Num ID:Phone Num Name" separated by a ";" character
  // Structure: [
  //  ["User ID 1", "User Name 1", "Phone Num ID A:Phone Num Name A;Phone Num ID B..."]
  //  ["User ID 2", "User Name 2", "Phone Num ID G:Phone Num Name G;Phone Num ID H..."],
  //  ...
  // ]
  var new_userdb_data = [];
  var num = 0;

  // If the "UserDB" sheet is empty...
  if((userdb_data.length == 1) && (userdb_data[0].length == 1)) {
    SpreadsheetApp.getUi().alert('No users found in the "UserDB" sheet, skipping step.');
    return;

  // ... else, the "UserDB" sheet already contains data...
  } else {
    SpreadsheetApp.getUi().alert("NOTE: If you have more than 50 users, this process may take \n" +
      "a few minutes to avoid hitting Aircall's API max [Calls-per-Minute] limit.");
  }

  // For each Aircall user...
  for (var i = 0; i < userdb_data.length; i++) {

    // Get the User ID and Username
    var user_id = parseInt(userdb_data[i][0]).toFixed(0);
    var username = userdb_data[i][1];

    new_userdb_data.push([user_id, username, ""]);

    if (!isNaN(user_id) && user_id) {
      // Define the url for the Aircall User API call for this specific user
      var user_url = "https://api.aircall.io/v1/users/" + user_id;

      // Parse the call response
      var user_response = UrlFetchApp.fetch(user_url, get_params);
      var user_json = user_response.getContentText();
      var user_data = JSON.parse(user_json);

      // For each number this user is assigned to...
      for (var j = 0; j < user_data.user.numbers.length; j++) {
        // Get the Number ID and Number Name
        var number_id = user_data.user.numbers[j].id;
        var number_name = user_data.user.numbers[j].name;

        // Keep track of it for this specific user
        new_userdb_data[i][2] += (number_id + ":" + number_name + ";");
      }

      // Trim last character of the string that represents the list of
      // numbers this user is assigned to
      new_userdb_data[i][2] = new_userdb_data[i][2].substring(0, new_userdb_data[i][2].length-1);

      num++;

      // If we have reached the max number of API calls per minute, sleep for 1 minute...
      if (num == MAX_CALLS) {
        Logger.log("Sleeping for 1 minute to avoid hitting API call max limit...");
        num = 0;
        Utilities.sleep(60000);
      }
    }
  }

  // Write the Aircall user <-> Phone number assignments to the "UserDB" sheet
  userdb_sheet.getRange("A1:C" + userdb_data.length.toString()).setValues(new_userdb_data);
}
