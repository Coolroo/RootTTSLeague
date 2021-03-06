//Data offsets
var timeStampOffset = 0;
var nameOffset = 3;
var statsOffset = 5;
var dataIDOffset = 1;
var dataTeamOffset = 5;
var dataWinnerOffset = 10;
var dataPointsOffset = 11;
var dataMapOffset = 15; 
var dataLandmarkOffset = 16;
var validityOffset = 18;
var IDOffset = 3;
var unpickeedFactionoffset = 16;

//Faction Data offsets
var factionOffset = 9;
var vagabondOffset = 28;

//Columns
var playerColumns = 110;
var factionColumns = 50;

//Constants
const domTypes = ["fox", "bunny", "mouse", "bird"];
const maps = ["Summer", "Winter", "Lake", "Mountain"];
const landmarks = ["Black Market", "Lost City", "The Tower", "Elder Treetop", "Legendary Forge", "The Ferry", "No Landmark"];
const teams = ["Marquise de Cat", "Eyrie Dynasties", "Woodland Alliance", "Vagabond", "Riverfolk Company", "Lizard Cult", "Underground Duchy", "Corvid Conspiracy", "Lord of the Hundreds", "Keepers in Iron"];
const vagabonds = ["Adventurer", "Arbiter", "Harrier", "Ranger", "Ronan", "Scoundrel", "Thief", "Tinker", "Vagrant"];

var numPlayers = 4;

//Math Constants
var roundVal = 2;

//Season Constants
const seasons = [{"startDate": new Date("May 22, 2022 00:00:00"), "endDate": null}];

var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var dataSheet = spreadsheet.getSheetByName("Data Input");
var allData = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, dataSheet.getLastColumn()).getValues();


//Invalid Games
var falseGames = [108, 121];
var invalidGames = new Array(allData.length).fill(null);
for(var i = 0; i<invalidGames.length; i++)
{
  var newDict = {"invalid": false, "reason": ""};
  invalidGames[i] = newDict;
}
for(var i = 0; i<falseGames.length; i++)
{
  invalidGames[falseGames[i]]["invalid"] = true;
  invalidGames[falseGames[i]]["reason"] = "(Manual Override)";
}

/**
 * The formSubmitted function is called when a player submits a new game
 * It takes all of the data from the Data Input sheet and checks to see if it is valid.
 * Invalid data will not be processed
 *
 * @docauthor Trelent
 */
function formSubmitted() 
{

  //Load the spreadsheet and get the data
  spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  dataSheet = spreadsheet.getSheetByName("Data Input");
  allData = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, dataSheet.getLastColumn()).getValues();

  //Check the submissions for validity and mark invalid games
  getValidGames();

  //Remove invalid games from the data array
  var splices = 0;
  for(var i = 0; i<invalidGames.length; i++)
  {
    if(invalidGames[i]["invalid"])
    {
      console.log("GAME " + (i + 1) + " IS INVALID");
      allData.splice(i - splices++, 1);
    }
  }
 
  //Compile the stats with the spliced data
  var stats = compileStats();
  console.log("Compiled Data");

  //Refresh the spreadsheet 
  refreshPlayerStats(stats["PlayerStats"]);
  refreshPlayerDomStats(stats["PlayerStats"]);
  refreshFactionStats(stats["FactionStats"], stats["VagabondStats"]);
  refreshFactionDomStats(stats["DomStats"]);
  refreshFactionH2H(stats["FactionStats"]);
}

//Spreadsheet update functions

function refreshPlayerStats(playerStats)
{

  var statsSheet = spreadsheet.getSheetByName("ProcessedPlayerData");
  statsSheet.getRange(statsOffset,1,999, playerColumns).clearContent();
  var players = Object.keys(playerStats);
  var totalStats = [];
  for(var i = 0; i<players.length; i++)
  {
    var myStats = playerStats[players[i]];
    var gameStats = myStats["gameStats"];
    var teamStats = gameStats["teams"];
    var localPlayerStats = [players[i],gameStats["games"],gameStats["wins"],(gameStats["wins"]*100/gameStats["games"]).toFixed(roundVal)+"%",gameStats["dom"]["generalStats"]["general"]["games"],gameStats["dom"]["generalStats"]["general"]["wins"],gameStats["points"],(gameStats["games"] - (gameStats["dom"]["generalStats"]["general"]["games"] + gameStats["coalition"]["generalStats"]["games"]) <= 0) ? 0 : Math.round(gameStats["points"]/(gameStats["games"]-(gameStats["dom"]["generalStats"]["general"]["games"] + gameStats["coalition"]["generalStats"]["games"])))];

    for(var j = 0; j<numPlayers; j++)
    {
      var localTurnStats = gameStats["turnStats"][j];
      var thisArr = [localTurnStats["games"], localTurnStats["wins"], (localTurnStats["games"] == 0) ? "0%" : (localTurnStats["wins"]*100/localTurnStats["games"]).toFixed(roundVal) + "%"];
      for(var k = 0; k<thisArr.length; k++)
      {
        localPlayerStats.push(thisArr[k]);
      }
    }
    for(var j = 0; j<teams.length; j++)
    {
      var localTeamStats = teamStats[teams[j]];
      var thisArr = [localTeamStats["games"], localTeamStats["wins"], (localTeamStats["games"] == 0) ? "0%" : (localTeamStats["wins"]*100/localTeamStats["games"]).toFixed(roundVal) + "%"];
      for(var k = 0; k<thisArr.length; k++)
      {
        localPlayerStats.push(thisArr[k]);
      }
    }
    for(var j = 0; j<vagabonds.length; j++)
    {
      var localVagabondStats = teamStats["Vagabond"]["stats"][vagabonds[j]];
      var thisArr = [localVagabondStats["games"], localVagabondStats["wins"], (localVagabondStats["games"] == 0) ? "0%" : (100*localVagabondStats["wins"]/localVagabondStats["games"]).toFixed(roundVal) + "%"];
      for(var k = 0; k<thisArr.length; k++)
      {
        localPlayerStats.push(thisArr[k]);
      }
    }
    for(var j = 0; j<maps.length; j++)
    {
      var localMapStats = gameStats["maps"][maps[j]];
      var thisArr = [localMapStats["games"], localMapStats["wins"], (localMapStats["games"] == 0) ? "0%" : (100*localMapStats["wins"]/localMapStats["games"]).toFixed(roundVal) + "%"];
      for(var k = 0; k<thisArr.length; k++)
      {
        localPlayerStats.push(thisArr[k]);
      }
    }
    for(var j = 0; j<landmarks.length; j++)
    {
      var localLandmarkStats = gameStats["landmarks"][landmarks[j]];
      var thisArr = [localLandmarkStats["games"], localLandmarkStats["wins"], (localLandmarkStats["games"] == 0) ? "0%" : (100*localLandmarkStats["wins"]/localLandmarkStats["games"]).toFixed(roundVal) + "%"]
      for(var k = 0; k<thisArr.length; k++)
      {
        localPlayerStats.push(thisArr[k]);
      }
    }
    totalStats.push(localPlayerStats);
  }
  statsSheet.getRange(statsOffset,1,totalStats.length, playerColumns).setValues(totalStats);
  return playerStats;

}

function refreshFactionStats(teamStats, vagaStats)
{
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var factionSheet = spreadsheet.getSheetByName("Faction Stats");

  var finalData = []
  for(var i = 0; i<teams.length; i++)
  {
    var myData = teamStats[teams[i]];
    var myBaseStats = myData["baseStats"];
    var myTurnStats = myData["turnStats"];
    var myMapStats = myData["mapStats"];
    var myLandmarkStats = myData["landmarkStats"];
    var thisArr = [teams[i], myBaseStats["drafted"], myBaseStats["games"], myBaseStats["wins"], (myBaseStats["games"] == 0) ? "0%" : (100*myBaseStats["wins"]/myBaseStats["games"]).toFixed(roundVal) + "%"];
    for(var j = 0; j<4; j++)
    {
      var thisTurn = myTurnStats[j];
      var localArr = [thisTurn["games"], thisTurn["wins"], (thisTurn["games"] == 0) ? "0%" : (100*thisTurn["wins"]/thisTurn["games"]).toFixed(roundVal) + "%"];
      for(var k = 0; k<localArr.length; k++)
      {
        thisArr.push(localArr[k]);
      }
    }
    for(var j = 0; j<maps.length; j++)
    {
      var thisMap = myMapStats[maps[j]];
      var localArr = [thisMap["games"], thisMap["wins"], (thisMap["games"] == 0) ? "0%" : (100*thisMap["wins"]/thisMap["games"]).toFixed(roundVal) + "%"];
      for(var k = 0; k<localArr.length; k++)
      {
        thisArr.push(localArr[k]);
      }
    }
    for(var j = 0; j<landmarks.length; j++)
    {
      var thisLandmark = myLandmarkStats[landmarks[j]];
      var localArr = [thisLandmark["games"], thisLandmark["wins"], (thisLandmark["games"] == 0) ? "0%" : (100*thisLandmark["wins"]/thisLandmark["games"]).toFixed(roundVal) + "%"];
      for(var k = 0; k<localArr.length; k++)
      {
        thisArr.push(localArr[k]);
      }
    }
    finalData.push(thisArr);
  } 

  var finalVagaData = []
  for(var i = 0; i<vagabonds.length; i++)
  {
    var myData = vagaStats[vagabonds[i]];
    var myBaseStats = myData["baseStats"];
    var myTurnStats = myData["turnStats"];
    var myMapStats = myData["mapStats"];
    var myLandmarkStats = myData["landmarkStats"];
    var thisArr = [vagabonds[i], myBaseStats["drafted"], myBaseStats["games"], myBaseStats["wins"], (myBaseStats["games"] == 0) ? "0%" : (100*myBaseStats["wins"]/myBaseStats["games"]).toFixed(roundVal) + "%"];
    for(var j = 0; j<4; j++)
    {
      var thisTurn = myTurnStats[j];
      var localArr = [thisTurn["games"], thisTurn["wins"], (thisTurn["games"] == 0) ? "0%" : (100*thisTurn["wins"]/thisTurn["games"]).toFixed(roundVal) + "%"];
      for(var k = 0; k<localArr.length; k++)
      {
        thisArr.push(localArr[k]);
      }
    }
    for(var j = 0; j<maps.length; j++)
    {
      var thisMap = myMapStats[maps[j]];
      var localArr = [thisMap["games"], thisMap["wins"], (thisMap["games"] == 0) ? "0%" : (100*thisMap["wins"]/thisMap["games"]).toFixed(roundVal) + "%"];
      for(var k = 0; k<localArr.length; k++)
      {
        thisArr.push(localArr[k]);
      }
    }
    for(var j = 0; j<landmarks.length; j++)
    {
      var thisLandmark = myLandmarkStats[landmarks[j]];
      var localArr = [thisLandmark["games"], thisLandmark["wins"], (thisLandmark["games"] == 0) ? "0%" : (100*thisLandmark["wins"]/thisLandmark["games"]).toFixed(roundVal) + "%"];
      for(var k = 0; k<localArr.length; k++)
      {
        thisArr.push(localArr[k]);
      }
    }
    finalVagaData.push(thisArr);
  } 
  factionSheet.getRange(factionOffset, 1, teams.length, factionColumns).setValues(finalData);
  factionSheet.getRange(vagabondOffset, 1, vagabonds.length, factionColumns).setValues(finalVagaData);
  return {"teamStats": teamStats, "vagaStats": vagaStats};
}

function refreshFactionDomStats(domStats)
{
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var domSheet = spreadsheet.getSheetByName("Dominance Stats");

  //Suits Template {"general": {"games": 0, "wins": 0}, "fox":{"games": 0, "wins": 0}, "bunny":{"games": 0, "wins": 0}, "mouse":{"games": 0, "wins": 0}, "bird":{"games": 0, "wins": 0}}


  var rowOffset = 12;
  var columns = 136;

  var finalData = [];
  for(var i = 0; i<teams.length; i++)
  {
    var team = teams[i];
    var myStats = domStats[team];
    var myMapStats = myStats["mapStats"];
    var myBaseStats = myStats["baseStats"];
    var myTurnStats = myStats["turnStats"];
    var myArr = [team, myBaseStats["general"]["games"], myBaseStats["general"]["wins"], (myBaseStats["general"]["games"] != 0) ? (100*myBaseStats["general"]["wins"]/myBaseStats["general"]["games"]).toFixed(roundVal) + "%" : "0%"];

    for(var j = 0; j<domTypes.length; j++)
    {
      var generalSuitStats = myBaseStats[domTypes[j]];
      var localArray = [generalSuitStats["games"], generalSuitStats["wins"], (generalSuitStats["games"] != 0) ? (100*generalSuitStats["wins"]/generalSuitStats["games"]).toFixed(roundVal) + "%" : "0%"];
      for(var k = 0; k<localArray.length; k++)
      {
        myArr.push(localArray[k]);
      }
    }

    for(var j = 0; j<4; j++)
    {
      var localTurnStats = myTurnStats[j];
      var generalTurnStats = localTurnStats["general"];
      var localArr = [generalTurnStats["games"], generalTurnStats["wins"], (generalTurnStats["games"] != 0) ? (100*generalTurnStats["wins"]/generalTurnStats["games"]).toFixed(roundVal) + "%" : "0%"];
      for(var k = 0; k<domTypes.length; k++)
      {
        var domType = domTypes[k];
        var domSuitStats = localTurnStats[domType];
        localArr.push(domSuitStats["games"]);
        localArr.push(domSuitStats["wins"]);
        localArr.push((domSuitStats["games"] != 0) ? (100*domSuitStats["wins"]/domSuitStats["games"]).toFixed(roundVal) + "%" : "0%");
      }
      for(var k = 0; k<localArr.length; k++)
      {
        myArr.push(localArr[k]);
      }
    }
    for(var j = 0; j<maps.length; j++)
    {
      var localMapStats = myMapStats[maps[j]];
      var generalMapStats = localMapStats["general"];
      var localArr = [generalMapStats["games"], generalMapStats["wins"], (generalMapStats["games"] != 0) ? (100*generalMapStats["wins"]/generalMapStats["games"]).toFixed(roundVal) + "%" : "0%"];
      for(var k = 0; k<domTypes.length; k++)
      {
        var domType = domTypes[k];
        var domSuitStats = localMapStats[domType];
        localArr.push(domSuitStats["games"]);
        localArr.push(domSuitStats["wins"]);
        localArr.push((domSuitStats["games"] != 0) ? (100*domSuitStats["wins"]/domSuitStats["games"]).toFixed(roundVal) + "%" : "0%");
      }
      for(var k = 0; k<localArr.length; k++)
      {
        myArr.push(localArr[k]);
      }
    }
    finalData.push(myArr);
  }
  
  domSheet.getRange(rowOffset, 1, finalData.length, columns).setValues(finalData);
}

function refreshPlayerDomStats(playerStats)
{
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var domSheet = spreadsheet.getSheetByName("ProcessedPlayerDomData");
  var domSheetOffset = 12;
  var players = Object.keys(playerStats);
  var finalData = [];
  for(var i = 0; i<players.length; i++)
  {
    var player = players[i];
    var myDomStats = playerStats[player]["gameStats"]["dom"];
    var generalStats = myDomStats["generalStats"];
    var turnStats = myDomStats["turnStats"];
    var mapStats = myDomStats["maps"];

    var finalArr = [player, generalStats["general"]["games"], generalStats["general"]["wins"], (generalStats["general"]["games"] != 0) ? (100*generalStats["general"]["wins"]/generalStats["general"]["games"]).toFixed(roundVal) + "%" : "0%"];
    for(var j = 0; j<domTypes.length; j++) //Submit General Stats
    {
      var typeStats = generalStats[domTypes[j]];
      var myArr = [typeStats["games"], typeStats["wins"], (typeStats["games"] != 0) ? (100*typeStats["wins"]/typeStats["games"]).toFixed(roundVal) + "%" : "0%"];
      for(var k = 0; k<myArr.length; k++)
      {
        finalArr.push(myArr[k]);
      }
    }
    for(var j = 0; j<4; j++) //Submit Turn Stats
    {
      var typeStats = turnStats[j];
      var myArr = [typeStats["general"]["games"], typeStats["general"]["wins"], (typeStats["general"]["games"] != 0) ? (100*typeStats["general"]["wins"]/typeStats["general"]["games"]).toFixed(roundVal) + "%" : "0%"];
      for(var k = 0; k<domTypes.length; k++)
      {
        var domTypeStats = typeStats[domTypes[k]];
        var typeArr = [domTypeStats["games"], domTypeStats["wins"], (domTypeStats["games"] != 0) ? (100*domTypeStats["wins"]/domTypeStats["games"]).toFixed(roundVal) + "%" : "0%"];
        for(var l = 0; l<typeArr.length; l++)
        {
          myArr.push(typeArr[l]);
        }
      }
      for(var k = 0; k<myArr.length; k++)
      {
        finalArr.push(myArr[k]);
      }
    }
    for(var j = 0; j<maps.length; j++) //Submit Map Stats
    {
      var typeStats = mapStats[maps[j]];
      var myArr = [typeStats["general"]["games"], typeStats["general"]["wins"], (typeStats["general"]["games"] != 0) ? (100*typeStats["general"]["wins"]/typeStats["general"]["games"]).toFixed(roundVal) + "%" : "0%"];
      for(var k = 0; k<domTypes.length; k++)
      {
        var domTypeStats = typeStats[domTypes[k]];
        var typeArr = [domTypeStats["games"], domTypeStats["wins"], (domTypeStats["games"] != 0) ? (100*domTypeStats["wins"]/domTypeStats["games"]).toFixed(roundVal) + "%" : "0%"];
        for(var l = 0; l<typeArr.length; l++)
        {
          myArr.push(typeArr[l]);
        }
      }
      for(var k = 0; k<myArr.length; k++)
      {
        finalArr.push(myArr[k]);
      }
    }
    finalData.push(finalArr);
  }
  domSheet.getRange(domSheetOffset, 1, finalData.length, domSheet.getLastColumn()).setValues(finalData);
}

function refreshFactionH2H(stats)
{
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var factionSheet = spreadsheet.getSheetByName("Faction Head to Head Data");

  //General Data
  var generalRow = 6;
  var generalCol = 2;
  var finalGeneralData = [];
  for(var i = 0; i<teams.length+1; i++)
  {
    finalGeneralData.push([]);
  }
  var firstRow = ["Factions"];
  for(var i = 0; i<teams.length; i++)
  {
    firstRow.push(teams[i]);
  }
  finalGeneralData[0] = firstRow;

  for(var i = 0; i<teams.length; i++)
  {
    var thisArr = []
    thisArr.push(teams[i]);
    //console.log(stats[teams[i]]);
    for(var j = 0; j<teams.length; j++)
    {
      if(i==j)
      {
        thisArr.push('');
        continue;
      }
      var ratio = (stats[teams[i]]["H2H"][teams[j]]["games"] > 0) ? (100*stats[teams[i]]["H2H"][teams[j]]["wins"]/stats[teams[i]]["H2H"][teams[j]]["games"]).toFixed(roundVal) + "%" : "N/A";
      thisArr.push(ratio);
    }
    finalGeneralData[i + 1] = thisArr;
  }
  console.log("final Data Length = " + finalGeneralData.length );
  factionSheet.getRange(generalRow, generalCol, teams.length+1, teams.length+1).setValues(finalGeneralData);

  //Win/Loss Data
  var winRow = 6;
  var winCol = 14;
  var finalWinData = [];
  for(var i = 0; i<teams.length+1; i++)
  {
    finalWinData.push([]);
  }
  var firstRow = ["Factions"];
  for(var i = 0; i<teams.length; i++)
  {
    firstRow.push(teams[i]);
  }
  finalWinData[0] = firstRow;

  for(var i = 0; i<teams.length; i++)
  {
    var thisArr = []
    thisArr.push(teams[i]);
    for(var j = 0; j<teams.length; j++)
    {
      if(i==j)
      {
        thisArr.push('');
        continue;
      }
      var ratio = (stats[teams[i]]["H2H"][teams[j]]["losses"] > 0) ? (100*stats[teams[i]]["H2H"][teams[j]]["wins"]/stats[teams[i]]["H2H"][teams[j]]["losses"]).toFixed(roundVal) + "%" : "N/A";
      thisArr.push(ratio);
    }
    finalWinData[i + 1] = thisArr;
  }
  console.log("final Data Length = " + finalGeneralData.length );
  factionSheet.getRange(winRow, winCol, teams.length+1, teams.length+1).setValues(finalWinData);
}

//Data Processing Functions

function compileStats()
{
  console.log("Compiling Stats");

  //Variables & Constants
  var stats = {"PlayerStats": {}, "FactionStats": {}, "VagabondStats": {}, "DomStats": {}, "SeasonStats": {}};
  const today = new Date();

  //initialize the stat list for overall stats, and all seasons
  var initStats = [stats];
  for(var i = 0; i<seasons.length; i++)
  {
    initStats.push(stats["SeasonStats"][seasons[i]] = {"PlayerStats": {}, "FactionStats": {}, "VagabondStats": {}, "DomStats": {}});
  }
  
  //Create stats for all seasons/overall to store the data
  for(var l = 0; l<initStats.length; l++)
  {
      //Personal stat variables
      var playerStats = initStats[l]["PlayerStats"];
      var teamStats = initStats[l]["FactionStats"];
      var vagaStats = initStats[l]["VagabondStats"];
      var domStats = initStats[l]["DomStats"];

      //Initialize the stats for all factions
      for(var i = 0; i<teams.length; i++)
      {
        console.log("Team (" + teams[i] + ") Initialized");
        teamStats[teams[i]] = {
          "H2H": {}, 
          "turnStats":[{"games":0, "wins":0}, {"games":0, "wins":0}, {"games":0, "wins":0}, {"games":0, "wins":0}], 
          "mapStats":{"Summer":{"games":0, "wins":0}, "Winter":{"games":0, "wins":0}, "Lake":{"games":0, "wins":0}, "Mountain":{"games":0, "wins":0}},
          "baseStats":{"games":0, "wins":0, "drafted":0}, 
          "landmarkStats": {}
      };
        
        //Initialize all landmark stats
        for(var j = 0; j<landmarks.length; j++)
        {
          teamStats[teams[i]]["landmarkStats"][landmarks[j]] = {"games":0, "wins": 0};
        }

        //Initialize Head-to-Head data
        for(var j = 0; j<teams.length; j++)
        {
          teamStats[teams[i]]["H2H"][teams[j]] = {"losses": 0, "wins": 0, "games": 0};
        }
      }

      //Vagabond Init
      for(var i = 0; i<vagabonds.length; i++)
      {
        vagaStats[vagabonds[i]] = {
          "turnStats":[{"games":0, "wins":0},{"games":0, "wins":0},{"games":0, "wins":0},{"games":0, "wins":0}], 
          "mapStats":{"Summer":{"games":0, "wins":0},"Winter":{"games":0, "wins":0},"Lake":{"games":0, "wins":0},"Mountain":{"games":0, "wins":0}}, 
          "baseStats":{"games":0, "wins":0, "drafted":0}, 
          "landmarkStats": {}
        };

        //Initialize Landmark Stats
        for(var j = 0; j<landmarks.length; j++)
        {
          vagaStats[vagabonds[i]]["landmarkStats"][landmarks[j]] = {"games": 0, "wins": 0};
        }
      }

      //Dom Init
      for(var i = 0; i<teams.length; i++)
      {
        var domTeamStats = {"baseStats": {}, "turnStats": [{},{},{},{}], "mapStats": {}};
        domTeamStats["baseStats"] = {
          "general": {"games": 0, "wins": 0}, 
          "fox":{"games": 0, "wins": 0}, 
          "bunny":{"games": 0, "wins": 0}, 
          "mouse":{"games": 0, "wins": 0}, 
          "bird":{"games": 0, "wins": 0}
        };

        //Initialize turn based dom stats
        for(var j = 0; j<4; j++)
        {
          domTeamStats["turnStats"][j] = {"general": {"games": 0, "wins": 0}, "fox":{"games": 0, "wins": 0}, "bunny":{"games": 0, "wins": 0}, "mouse":{"games": 0, "wins": 0}, "bird":{"games": 0, "wins": 0}};
          domTeamStats["mapStats"][maps[j]] = {"general": {"games": 0, "wins": 0}, "fox":{"games": 0, "wins": 0}, "bunny":{"games": 0, "wins": 0}, "mouse":{"games": 0, "wins": 0}, "bird":{"games": 0, "wins": 0}};
        }
        domStats[teams[i]] = domTeamStats;
      }
  }
  console.log("PROCESSING ALL DATA");

  //Process all valid games
  for(var i = 0; i<allData.length; i++)
  {

    //Collect data vars
    var data = allData[i];
    var timeStamp = data[timeStampOffset];
    var playerIDs = [data[dataIDOffset + 0], data[dataIDOffset + 1], data[dataIDOffset + 2], data[dataIDOffset + 3]];
    var playerTeams = [data[dataTeamOffset + 0], data[dataTeamOffset + 1], data[dataTeamOffset + 2], data[dataTeamOffset + 3]];
    var winningPlayer = getWinner(data[dataWinnerOffset].toString());
    var playerPoints = [data[dataPointsOffset + 0], data[dataPointsOffset + 1], data[dataPointsOffset + 2], data[dataPointsOffset + 3]];
    var map = data[dataMapOffset];
    var landmark = data[dataLandmarkOffset];


    //Get season this game falls under
    var validSeasons = [];
    for(var j = 0; j<seasons.length; j++)
    {
      const thisSeason = seasons[j];
      const thisDay = new Date(timeStamp);
      if(thisSeason["startDate"].getTime() - thisDay.getTime() < 0 && (thisSeason["endDate"] == null || thisSeason["endDate"].getTime() - thisDay.getTime() > 0))
      {
        validSeasons.push(thisSeason);
      }
    }

    //Determine the winner
    var winIndexes = new Array(numPlayers).fill(false);
    winIndexes[winningPlayer] = true;

    //Get all the valid seasons needing to be processed
    var statsList = [stats];
    for(var vs = 0; vs<validSeasons.length; vs++)
    {
      statsList.push(stats["SeasonStats"][seasons[vs]]);
    }
    
    //For each season needing to be processed
    for(var stat = 0; stat<statsList.length; stat++)
    {
      //Local variables for this stat collection
      var playerStats = statsList[stat]["PlayerStats"];
      var teamStats = statsList[stat]["FactionStats"];
      var vagaStats = statsList[stat]["VagabondStats"];
      var domStats = statsList[stat]["DomStats"];

      //Collect player stats for this game
      for(var j = 0; j<numPlayers; j++)
      {
        //Get the players in this game & create stats for this player if they have not played a game in this season
        var players = Object.keys(playerStats);
        if(players.indexOf(playerIDs[j]) < 0)
        {
          playerStats[playerIDs[j]] = createPlayerStats(numPlayers);
        }
        
        //Get the player stats and determine if they won (winner = 1 === WIN, winner = 0 === LOSE)
        var myStats = playerStats[playerIDs[j]]["gameStats"];
        var winner = (winIndexes[j]) ? 1 : 0;

        //Update win stats
        myStats["games"]++;
        myStats["wins"] += winner;


        //Update Dom Stats
        if(isDom(playerPoints[j].toString()))
        {
          var playerDomStats = myStats["dom"];
          var domSuit = getDomType(playerPoints[j].toString());

          playerDomStats["generalStats"]["general"]["games"]++;
          playerDomStats["generalStats"]["general"]["wins"]+=winner;

          playerDomStats["generalStats"][domSuit]["games"]++;
          playerDomStats["generalStats"][domSuit]["wins"]+=winner;

          playerDomStats["maps"][map]["general"]["games"]++;
          playerDomStats["maps"][map]["general"]["wins"] += winner;

          playerDomStats["maps"][map][domSuit]["games"]++;
          playerDomStats["maps"][map][domSuit]["wins"] += winner;

          var domTurnStats = playerDomStats["turnStats"][j];

          domTurnStats["general"]["games"]++;
          domTurnStats["general"]["wins"]+= winner;

          domTurnStats[domSuit]["games"]++;
          domTurnStats[domSuit]["wins"] += winner;
          
        }

        //Does not happen anymore as coalitions are banned, but leaving here in case coalitions return
        else if(isCoalition(playerPoints[j].toString()))
        {
          var coalitionStats = myStats["coalition"];

          coalitionStats["generalStats"]["games"]++;
          coalitionStats["generalStats"]["wins"]+=winner;

          coalitionStats["maps"][map]["games"]++;
          coalitionStats["maps"][map]["wins"] += winner;


          var coalitionTurnStats = coalitionStats["turnStats"][j];

          coalitionTurnStats["games"]++;
          coalitionTurnStats["wins"]+= winner;

        }
        else
        {
          //Otherwise this is a non-dom game, so track the player's points
          myStats["points"] += Number(playerPoints[j]);
        }
        

        //Update player stats based on Turn Order, Map, & Landmarks

        myStats["turnStats"][j]["games"]++;
        myStats["turnStats"][j]["wins"] += winner;

        myStats["maps"][map]["games"]++;
        myStats["maps"][map]["wins"] += winner;
        
        myStats["landmarks"][landmark]["games"]++;
        myStats["landmarks"][landmark]["wins"] += winner;

        //Find the team of this player and update it's stats
        var myTeams = myStats["teams"];
        switch(playerTeams[j])
        {
          case "Marquise de Cat":
          myTeams[teams[0]]["games"]++;
          myTeams[teams[0]]["wins"]+= winner;
          break;

          case "Eyrie Dynasties":
          myTeams[teams[1]]["games"]++;
          myTeams[teams[1]]["wins"]+= winner;
          break;

          case "Woodland Alliance":
          myTeams[teams[2]]["games"]++;
          myTeams[teams[2]]["wins"]+= winner;
          break;

          case "Riverfolk Company":
          myTeams[teams[4]]["games"]++;
          myTeams[teams[4]]["wins"]+= winner;
          break;

          case "Lizard Cult":
          myTeams[teams[5]]["games"]++;
          myTeams[teams[5]]["wins"]+= winner;
          break;

          case "Underground Duchy":
          myTeams[teams[6]]["games"]++;
          myTeams[teams[6]]["wins"]+= winner;
          break;

          case "Corvid Conspiracy":
          myTeams[teams[7]]["games"]++;
          myTeams[teams[7]]["wins"]+= winner;
          break;

          case "Lord of the Hundreds":
          myTeams[teams[8]]["games"]++;
          myTeams[teams[8]]["wins"]+= winner;
          break;

          case "Keepers in Iron":
          myTeams[teams[9]]["games"]++;
          myTeams[teams[9]]["wins"]+= winner;
          break;
          
          default:
          myTeams[teams[3]]["games"]++;
          myTeams[teams[3]]["wins"]+= winner;
          var vagaName = playerTeams[j].substring(9);
          var vagaTeams = myTeams[teams[3]]["stats"];
          switch(vagaName)
          {
            case "(Adventurer)":
            vagaTeams[vagabonds[0]]["games"]++;
            vagaTeams[vagabonds[0]]["wins"] += winner;
            break;

            case "(Arbiter)":
            vagaTeams[vagabonds[1]]["games"]++;
            vagaTeams[vagabonds[1]]["wins"] += winner;
            break;

            case "(Harrier)":
            vagaTeams[vagabonds[2]]["games"]++;
            vagaTeams[vagabonds[2]]["wins"] += winner;
            break;

            case "(Ranger)":
            vagaTeams[vagabonds[3]]["games"]++;
            vagaTeams[vagabonds[3]]["wins"] += winner;
            break;

            case "(Ronan)":
            vagaTeams[vagabonds[4]]["games"]++;
            vagaTeams[vagabonds[4]]["wins"] += winner;
            break;

            case "(Scoundrel)":
            vagaTeams[vagabonds[5]]["games"]++;
            vagaTeams[vagabonds[5]]["wins"] += winner;
            break;

            case "(Thief)":
            vagaTeams[vagabonds[6]]["games"]++;
            vagaTeams[vagabonds[6]]["wins"] += winner;
            break;

            case "(Tinker)":
            vagaTeams[vagabonds[7]]["games"]++;
            vagaTeams[vagabonds[7]]["wins"] += winner;
            break;

            case "(Vagrant)":
            vagaTeams[vagabonds[8]]["games"]++;
            vagaTeams[vagabonds[8]]["wins"] += winner;
            break;
          }
          break;
        }
      }

      //Compile Faction & Vagabond Stats

      //Game Variables
      var playingTeams = [];
      var winningPlayer = getWinner(allData[i][dataWinnerOffset]);
      var winningTeam = "";

      var seenVaga = false;

      //Check the team of each player in the game IN REVERSE DRAFT ORDER
      for(var c = numPlayers; c>=0; c--)
      {
        //Check undrafted last
        var j = 0;
        if(c == 0)
        {
          j = numPlayers;
        }
        else
        {
          j = c - 1;
        }
        //Local player variables
        var won = (winningPlayer == j) ? 1 : 0;
        var map = allData[i][dataMapOffset];
        var landmark = allData[i][dataLandmarkOffset];
        var team = allData[i][dataTeamOffset+j];
        var myTeamStats = null;
        var isVaga = false;
        var localTeam = "";

        //Find the player's team
        switch(team)
        {
          case "Marquise de Cat":
          localTeam = teams[0];
          myTeamStats = teamStats[teams[0]];
          break;

          case "Eyrie Dynasties":
          localTeam = teams[1];
          myTeamStats = teamStats[teams[1]];
          break;

          case "Woodland Alliance":
          localTeam = teams[2];
          myTeamStats = teamStats[teams[2]];
          break;

          case "Riverfolk Company":
          localTeam = teams[4];
          myTeamStats = teamStats[teams[4]];
          break;

          case "Lizard Cult":
          localTeam = teams[5];
          myTeamStats = teamStats[teams[5]];
          break;

          case "Underground Duchy":
          localTeam = teams[6];
          myTeamStats = teamStats[teams[6]];
          break;

          case "Corvid Conspiracy":
          localTeam = teams[7];
          myTeamStats = teamStats[teams[7]];
          break;

          case "Lord of the Hundreds":
          localTeam = teams[8];
          myTeamStats = teamStats[teams[8]];
          break;

          case "Keepers in Iron":
          localTeam = teams[9];
          myTeamStats = teamStats[teams[9]];
          break;
          
          default:
          localTeam = teams[3];
          myTeamStats = teamStats[teams[3]];
          isVaga = true;
          break;
        }

        //Get the Faction's stats
        var myBaseStats = myTeamStats["baseStats"];
        var myTurnStats = myTeamStats["turnStats"];
        var myMapStats = myTeamStats["mapStats"];
        var myLandmarkStats = myTeamStats["landmarkStats"];

        //Mark this faction as drafted (Do not count second vagabond)
        if(!(isVaga && seenVaga))
        {
          myBaseStats["drafted"]++;
        }
        

        seenVaga = seenVaga || isVaga;
        //If we are looking at a faction that was chosen
        if(j != numPlayers)
        {
          //If we won, then set us as the winning team
          if(won == 1)
          {
            winningTeam = localTeam;
          }

          //Mark this team as a team that was in the game
          playingTeams.push(localTeam);

          //Update the faction stats
          myBaseStats["games"]++;
          myBaseStats["wins"] += won;

          myTurnStats[j]["games"]++;
          myTurnStats[j]["wins"] += won;

          myMapStats[map]["games"]++;
          myMapStats[map]["wins"] += won;

          myLandmarkStats[landmark]["games"]++;
          myLandmarkStats[landmark]["wins"] += won;
        }
        //If this player was a vagabond, find out which vagabond they were
        if(isVaga)
        {
          var vagaName = team.substring(9);
          var myVaga = null;
          switch(vagaName)
          {
            case "(Adventurer)":
            myVaga = vagabonds[0];
            break;

            case "(Arbiter)":
            myVaga = vagabonds[1];
            break;

            case "(Harrier)":
            myVaga = vagabonds[2];
            break;

            case "(Ranger)":
            myVaga = vagabonds[3];
            break;

            case "(Ronan)":
            myVaga = vagabonds[4];
            break;

            case "(Scoundrel)":
            myVaga = vagabonds[5];
            break;

            case "(Thief)":
            myVaga = vagabonds[6];
            break;

            case "(Tinker)":
            myVaga = vagabonds[7];
            break;

            case "(Vagrant)":
            myVaga = vagabonds[8];
            break;
          }

          //Collect local Vagabond Stats
          myBaseStats = vagaStats[myVaga]["baseStats"];
          myTurnStats = vagaStats[myVaga]["turnStats"];
          myMapStats = vagaStats[myVaga]["mapStats"];
          myLandmarkStats = vagaStats[myVaga]["landmarkStats"];
          
          //Check to see if another vagabond was drafted as to not fuck up the stats
          myBaseStats["drafted"]++;

          //If this player was chosen
          if(j != numPlayers)
          {
            myBaseStats["games"]++;
            myBaseStats["wins"] += won;

            myTurnStats[j]["games"]++;
            myTurnStats[j]["wins"] += won;

            myMapStats[map]["games"]++;
            myMapStats[map]["wins"] += won;
            
            myLandmarkStats[landmark]["games"]++;
            myLandmarkStats[landmark]["wins"] += won;
          }
        }
      }
      for(var j = 0; j<playingTeams.length; j++)
      {
        teamStats[playingTeams[j]]["H2H"][winningTeam]["losses"]++;
        teamStats[winningTeam]["H2H"][playingTeams[j]]["wins"]++;
        for(var k = 0; k<playingTeams.length; k++)
        {
          teamStats[playingTeams[j]]["H2H"][playingTeams[k]]["games"]++;
        }
      }

      //Dom Stats
      var winningPlayer = getWinner(allData[i][dataWinnerOffset]);
      for(var j = 0; j<numPlayers; j++)
      {
        //Check if this player played a dom
        var points = allData[i][dataPointsOffset+j];
        if(isDom(points.toString()))
        {
          //Get the dom type and other game vars
          var domType = getDomType(points.toString());
          var won = (winningPlayer == j) ? 1 : 0;
          var map = allData[i][dataMapOffset];
          var team = allData[i][dataTeamOffset+j];
          
          //Collect dom stats
          var myStats = domStats[team];
          var myMapStats = myStats["mapStats"];
          var myBaseStats = myStats["baseStats"];
          var myTurnStats = myStats["turnStats"];

          myBaseStats["general"]["games"]++;
          myBaseStats["general"]["wins"] += won;

          myBaseStats[domType]["games"]++;
          myBaseStats[domType]["wins"] += won;

          myMapStats[map]["general"]["games"]++;
          myMapStats[map]["general"]["wins"] += won;

          myMapStats[map][domType]["games"]++;
          myMapStats[map][domType]["wins"] += won;

          myTurnStats[j]["general"]["games"]++;
          myTurnStats[j]["general"]["wins"] += won;

          myTurnStats[j][domType]["games"]++;
          myTurnStats[j][domType]["wins"] += won;
        }
        //Not used
        else if(isCoalition(points.toString())){}
      }
    }
    //Report compilation progress
    console.log(((i+1)*100/allData.length).toFixed(2) + "% Complete");
  }

  return stats;

}

function getValidGames()
{

  for(var i = 0; i<invalidGames.length; i++)
  {
    var data = allData[i];
    var playerTeams = [data[dataTeamOffset + 0], data[dataTeamOffset + 1], data[dataTeamOffset + 2], data[dataTeamOffset + 3]];
    var winningPlayer = getWinner(data[dataWinnerOffset].toString());
    var playerPoints = [data[dataPointsOffset + 0], data[dataPointsOffset + 1], data[dataPointsOffset + 2], data[dataPointsOffset + 3]];
    var playerIDs = [data[dataIDOffset + 0], data[dataIDOffset + 1], data[dataIDOffset + 2], data[dataIDOffset + 3]];

    for(var j = 0; j<playerTeams;j++)
    {
      if(isVagabond(playerTeams[j].toString()) && isDom(playerPoints[j].toString()))
      {
        invalidGames[i]["invalid"] = true;
        invalidGames[i]["reason"] += "(Vagabond Dom)";
        break;
      }
    }
    if(playerPoints[winningPlayer] < 30 && !isDom(playerPoints[winningPlayer].toString()))
    {
      invalidGames[i]["invalid"] = true;
      invalidGames[i]["reason"] += "(Invalid Winner Points)";
    }
    if(hasDuplicate(playerIDs))
    {
      invalidGames[i]["invalid"] = true;
      invalidGames[i]["reason"] += "(Duplicate Player ID)";
    }
    var values = [[]];
    values[0].push(invalidGames[i]["invalid"] ? "INVALID" : "VALID");
    values[0].push(invalidGames[i]["reason"]);
    dataSheet.getRange(i + 2, validityOffset + 1, 1, 2).setValues(values);
  }

}

//Data Helper Functions

function createPlayerStats(numPlayers)
{
  playerStats = {};
  playerStats["gameStats"] = {};
  playerStats["gameStats"]["turnStats"] = [];
  playerStats["gameStats"]["games"] = 0;
  playerStats["gameStats"]["dom"] = {"maps": {}, "turnStats": [], "generalStats": {"general": {"games":0, "wins":0}, "fox": {"games":0, "wins":0}, "bunny": {"games":0, "wins":0}, "mouse": {"games":0, "wins":0}, "bird": {"games":0, "wins":0}}};
  playerStats["gameStats"]["coalition"] = {"maps": {}, "turnStats": [], "generalStats": {"games": 0, "wins": 0}};
  playerStats["gameStats"]["wins"] = 0;
  playerStats["gameStats"]["points"] = 0;
  playerStats["gameStats"]["teams"] = {};
  playerStats["gameStats"]["maps"] = {};
  playerStats["gameStats"]["landmarks"] = {};
  for(var k = 0; k<numPlayers; k++)
  {
    playerStats["gameStats"]["turnStats"].push({"games":0, "wins":0});
    playerStats["gameStats"]["dom"]["turnStats"].push({"general": {"games": 0, "wins": 0}, "fox":{"games": 0, "wins": 0}, "bunny":{"games": 0, "wins": 0}, "mouse":{"games": 0, "wins": 0}, "bird":{"games": 0, "wins": 0}});
    playerStats["gameStats"]["coalition"]["turnStats"].push({"games":0, "wins":0});
  }
  for(var k = 0; k<maps.length; k++)
  {
    playerStats["gameStats"]["maps"][maps[k]] = {"games":0, "wins":0};
    playerStats["gameStats"]["dom"]["maps"][maps[k]] = {"general": {"games":0, "wins":0}, "fox": {"games":0, "wins":0}, "bunny": {"games":0, "wins":0}, "mouse": {"games":0, "wins":0}, "bird": {"games":0, "wins":0}};
    playerStats["gameStats"]["coalition"]["maps"][maps[k]] = {"games":0, "wins":0};
  }
  for(var k = 0; k<teams.length; k++)
  {
    playerStats["gameStats"]["teams"][teams[k]] = {"games":0, "wins":0};
  }
  for(var k = 0; k<landmarks.length; k++)
  {
    playerStats["gameStats"]["landmarks"][landmarks[k]] = {"games": 0, "wins": 0};
  }
  playerStats["gameStats"]["teams"]["Vagabond"]["stats"] = {};
  for(var k = 0; k<vagabonds.length; k++)
  {
    playerStats["gameStats"]["teams"]["Vagabond"]["stats"][vagabonds[k]] = {"games":0, "wins":0};
  }
  return playerStats;
}

function getWinner(text)
{
  if(RegExp("[fF]irst").test(text))
  {
    return 0;
  }
  else if(RegExp("[sS]econd").test(text))
  {
    return 1;
  }
  else if(RegExp("[tT]hird").test(text))
  {
    return 2;
  }
  else
  {
    return 3;
  }
}

function isDom(text)
{
  return RegExp("[dD]om").test(text);
}

function isCoalition(text)
{
  return RegExp("[cC]oalition").test(text);
}

function isVagabond(text)
{
  return RegExp("Vagabond").test(text);
}

function getCoalitionPlayer(text)
{
  return +text.substring(10,11);
}

function getDomType(text)
{
  if(RegExp("[fF]ox").test(text))
  {
    return "fox";
  }
  else if(RegExp("[bB]unny").test(text))
  {
    return "bunny";
  }
  else if(RegExp("[mM]ouse").test(text))
  {
    return "mouse";
  }
  else
  {
    return "bird";
  }
}

function hasDuplicate(arr)
{
  return (new Set(arr)).size != arr.length;
}
