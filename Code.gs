var MovieNum = 0;
var PrevNum = 1;
var UpDown = 2;
var MovieName = 3;
var Avg = 4;
var Scott = 5;
var Jess = 6;
var Brock = 7;
var Lauren = 8;
var Elaine = 9;
var Pat = 10;
var RatingWeekCol = 11;
var AvgOldSheet = 8;

var desc = {
    HOUSE_COLLINS: 1,
    HOUSE_RISOLUTE: 2,
    HOUSE_SCHRECK: 3,
    KUNGFU: 4,
    ACTION: 5,
    HORROR: 6,
    SCIFI: 7,
    ROMCOM: 8,
    XMAS: 9,
    BOOBS: 10,
    NOBOOBS: 11,
    SEQUEL: 12,
    NOT_RATED: 13,
    RATED_PG: 14,
    RATED_PG13: 15,
    RATED_TV14: 16,
    RATED_R: 17,
    R_1980s: 18,
    R_1990s: 19,
    R_2000s: 20,
    R_2010s: 21,
    LESS_90: 22,
    MORE_90: 23,
    ANIMAL_ATTACK: 24,
    PLACEHOLDER: 26,
    JESS_REMEMBER: 26,
    JESS_NO_REMEMBER: 27,
    SCOTT: 28,
    JESS: 29,
    PENIS: 30,
    LAUREN: 31,
    ELAINE: 32,
    PAT:33
}

var detailcols = {
    WEEK: 1,
    MOVIENAME: 2,
    RT: 4,
    IMDB: 5,
    HOUSE: 6,
    RATING: 7,
    RUNTIME: 8,
    BOXOFFICE: 9,
    YEAR: 10,
    BOOBS: 11,
    KUNGFU: 12,
    ACTION: 13,
    HORROR: 14,
    SCIFI: 15,
    ROMCOM: 16,
    XMAS: 17,
    SEQUEL: 19,
    ANIMAL: 20,
    JESS: 21,
    AVGRATING: 22
}

SpreadsheetApp.setActiveSpreadsheet(SpreadsheetApp.openById("1YVsTgDIsHHg08W6HgDV2-2HlijMFnTpwqTSG9n8QuMA"));
var ss = SpreadsheetApp.getActiveSpreadsheet();
var infosheet = ss.getSheetByName("MovieDetails");
var infodata = infosheet.getDataRange().getValues();

var ratingsheet = ss.getSheetByName("Latest");
var ratingdata = ratingsheet.getDataRange().getValues();

function refreshall(range, throwaway) { return range.length }

function Movie(week,title, imdb, house, boobs,karate,action,horror,sci,rom,xmas,seq,animal,jess)
{
  this.week = week;
  this.title = title;
  this.imdb = imdb;
  this.house = house;
  this.boobs = boobs;
  this.karate = karate;
  this.action = action;
  this.horror = horror;
  this.sci = sci;
  this,rom = rom;
  this.xmas = xmas;
  this.seq = seq;
  this.animal = animal;
  this.jess = jess;
}

function IndRankings(scott,jess,brock,lauren,elaine,pat)
{
  this.scott=scott;
  this.jess=jess;
  this.brock=brock;
  this.jess=jess;
  this.elaine=elaine;
  this.pat=pat;
}

function addNewMovie()
{
  var movie = new Movie(
    105,
    "Jaws", //Title
    "tt0073195", //Imdb
    "Risolute", //House
    "No", //Boobs
    "Yes", //Karate
    "No", //Action
    "Yes", //Horror
    "No", //SciFi/Fantasy
    "No", //RomCom
    "No", //Xmas
    "No", //Sequel
    "Yes", //AnimalAttack
    "Y" // Jess Remember?
  )

  var rankings = new IndRankings(
    104, //Scott
    104, //Jess
    104, //Brock
    104, //Lauren
    104, //Elaine
    104 //Pat
  )

  var m_prevWeek = movie.week - 1;

  var m_prevWeekSheet = ss.getSheetByName("Week "+ m_prevWeek);

  if (m_prevWeekSheet) {
    Logger.log(ss.getName());
    Logger.log("Week " + m_prevWeek + " sheet exists!");
    return -1;
  }

  // This copies 'Latest' to the previous 'Week #' sheet
  m_prevWeekSheet = ss.insertSheet("Week " + m_prevWeek);
  var m_latestSheet = ss.getSheetByName('Latest');
  var m_latestData = m_latestSheet.getDataRange();
  m_latestVals = m_latestData.getValues();
  m_prevWeekSheet.getRange(m_latestData.getA1Notation()).setValues(m_latestData.getValues());

  var sourceRange = m_latestSheet.getRange("A2:A200");
  var destRange=m_latestSheet.getRange("B2:B200"); 
  sourceRange.copyTo(destRange);

  // Add the movie line to movie details
  addMovieDetailRow(movie);
  pullMovieDetailsIntoSpreadsheet();

  // Add the new rankings line to the Latest sheet
  m_latestSheet.getRange(movie.week+1,1,1,12)//(start row, start column, number of rows, number of columns
   .setValues([[
               movie.week,
               "n/a",
               "",
               movie.title,
               movie.week,
               movie.week,
               movie.week,
               movie.week,
               movie.week,
               movie.week,
               movie.week,
               movie.week
              ]]);

  // Add each player's scores
  add_player_score_to_latest(movie.week, Brock, rankings.brock);

  // Add the new averages to the moviedetails sheet's latest week
  //addAverages();

  return;
}
function updateWeeks()
{
  var week = 83;
  var weekCol = 11;
  var avgCol = 0;
  var weekit=0;
  for(weekit=(ratingdata.length-1); weekit < ratingdata.length; weekit++)
  {
    if(weekit < 58)
    {
      avgCol = 8;
    }
    else
    {
      avgCol = 4;
    }


      var weeksheet_string = "Week "+ String(weekit);

    if(weekit == ratingdata.length-1)
    {
      weeksheet_string = "Latest";
    }

    var weeksheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(weeksheet_string);
    var weeksheetdata = weeksheet.getDataRange().getValues();
    Logger.log("Week " + weekit);
    for (i = 1; i < weeksheetdata.length; i++) {

      var ave = weeksheetdata[i][avgCol];
      var movieweek = weeksheetdata[i][11];

      Logger.log("Movie Week " + movieweek + " Avg: " + ave);
      var cell = infosheet.getRange(movieweek + 1, detailcols.AVGRATING+weekit);
      Logger.log("Setting cell ["+ (movieweek + 1) + "][" + (detailcols.AVGRATING+weekit) + "]");
      cell.setValue(ave);
    }

    week++;
  }
}

function homeFieldAdvantageByWeek(who_orig, week)
{
  var num_weeks = ratingdata.length - 1; // Length has a heading row
    //who_orig = 5;
  var who = who_orig;
  var avg_col = Avg;
  //week = 82;
  var diff = 0;
  if(week > num_weeks)
  {
    Logger.log("Invalid Week Number!");
  }
  else
  {
    if(week < 58)
    {
      who = who_orig - 3;
      avg_col = AvgOldSheet;
    }
    var weeksheet_string = "Week "+ String(week);
    if(week == num_weeks)
    {
      weeksheet_string = "Latest";
    }

    Logger.log("Getting homer for week " + weeksheet_string);
    var weeksheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(weeksheet_string);
    var weeksheetdata = weeksheet.getDataRange().getValues();

    Logger.log("Week " + week + " length is " + weeksheetdata.length);

    for (var i = 1; i < weeksheetdata.length; i++) {

        if (true == isWeekAHomePick(weeksheetdata[i][RatingWeekCol], who_orig)) {
            Logger.log("Before Movie: " + weeksheetdata[i][MovieName] + " Diff was " + diff);
            Logger.log("Subtracting " + weeksheetdata[i][who] + " and " + weeksheetdata[i][avg_col]);
            diff += weeksheetdata[i][who] - weeksheetdata[i][avg_col];
            Logger.log("After Movie: " + weeksheetdata[i][MovieName] + " Diff was " + diff);
        }
    }

  }

  return diff.toFixed(2) * -1;
}


function GetNumberOfMoviesByDescByWeek(descriptor, week) {
    var movielist = [];
    movielist = add_movies_by_desc(descriptor, week);
    return movielist.length;
}

function GetNumberOfMoviesByDesc(descriptor) {
    return GetNumberOfMoviesByDescByWeek(descriptor, ratingdata.length - 1);
}

function add_movies_by_desc(descriptor, week) {
    var movielist = [];

    Logger.log("Week is" + week + "look4desc " + descriptor);
    for (i = 1; i <= week; i++) {
        Logger.log("Movielist Length is " + movielist.length);
        if(true == isWeekValidForDesc(i, descriptor))
        {
          movielist.push(i);
        }
    }

    return movielist;
}

function isWeekValidForDesc(week, descriptor)
{
  var i = week;
  var result = false;
  if (infodata[i][detailcols.MOVIENAME] !== "")
  {
    switch (descriptor) {
        case desc.HOUSE_COLLINS:
            {
                if (infodata[i][detailcols.HOUSE] == "Collins") {
                    result = true;
                }
            }
            break;
        case desc.HOUSE_RISOLUTE:
            {
                if (infodata[i][detailcols.HOUSE] == "Risolute") {
                    result = true;
                }
            }
            break;
        case desc.HOUSE_SCHRECK:
            {
                if (infodata[i][detailcols.HOUSE] == "Schreck") {
                    result = true;
                }
            }
            break;
        case desc.KUNGFU:
            {
                if (infodata[i][detailcols.KUNGFU] == "Yes") {
                    result = true;
                }
            }
            break;
        case desc.ROMCOM:
            {
                if (infodata[i][detailcols.ROMCOM] == "Yes") {
                    result = true;
                }
            }
            break;
        case desc.HORROR:
            {
                if (infodata[i][detailcols.HORROR] == "Yes") {
                    result = true;
                }
            }
            break;
        case desc.XMAS:
            {
                if (infodata[i][detailcols.XMAS] == "Yes") {
                    result = true;
                }
            }
            break;
        case desc.SCIFI:
            {
                if (infodata[i][detailcols.SCIFI] == "Yes") {
                    result = true;
                }
            }
            break;
        case desc.ACTION:
            {
                if (infodata[i][detailcols.ACTION] == "Yes") {
                    result = true;
                }
            }
            break;
        case desc.BOOBS:
            {
                if (infodata[i][detailcols.BOOBS] == "Yes") {
                    result = true;
                }
            }
            break;
        case desc.NOBOOBS:
            {
                if (infodata[i][detailcols.BOOBS] == "No") {
                    result = true;
                }
            }
            break;
        case desc.SEQUEL:
            {
                if (infodata[i][detailcols.SEQUEL] == "Yes") {
                    result = true;
                }
            }
            break;
        case desc.NOT_RATED:
            {
                if (infodata[i][detailcols.RATING] == "Not Rated") {
                    result = true;
                }
                if (infodata[i][detailcols.RATING] == "Unrated") {
                    result = true;
                }
                if (infodata[i][detailcols.RATING] == "N/A") {
                    result = true;
                }
            }
            break;
        case desc.RATED_PG:
            {
                if (infodata[i][detailcols.RATING] == "PG") {
                    result = true;
                }
            }
            break;
        case desc.RATED_PG13:
            {
                if (infodata[i][detailcols.RATING] == "PG-13") {
                    result = true;
                }
            }
            break;
        case desc.RATED_TV14:
            {
                if (infodata[i][detailcols.RATING] == "TV-14") {
                    result = true;
                }
            }
            break;
        case desc.RATED_R:
            {
                if (infodata[i][detailcols.RATING] == "R") {
                    result = true;
                }
            }
            break;
        case desc.R_1980s:
            {
                var year = infodata[i][detailcols.YEAR];
                if ((year > 1979) && (year < 1990)) {
                    result = true;
                }
            }
            break;
        case desc.R_1990s:
            {
                var year = infodata[i][detailcols.YEAR];
                if ((year > 1989) && (year < 2000)) {
                    result = true;
                }
            }
            break;
        case desc.R_2000s:
            {
                var year = infodata[i][detailcols.YEAR];
                if ((year > 1999) && (year < 2010)) {
                    result = true;
                }
            }
            break;
        case desc.R_2010s:
            {
                var year = infodata[i][detailcols.YEAR];
                if ((year > 2009) && (year < 2020)) {
                    result = true;
                }
            }
            break;
        case desc.LESS_90:
            {
                var time = infodata[i][detailcols.RUNTIME];
                var array1 = [{}];

                array1 = time.split(" ");
                if (array1[0] < 91) {
                    result = true;
                }
            }
            break;
        case desc.MORE_90:
            {
                var time = infodata[i][detailcols.RUNTIME];
                var array1 = [{}];

                array1 = time.split(" ");
                if (array1[0] > 90) {
                    result = true;
                }
            }
            break;
        case desc.ANIMAL_ATTACK:
            {
                var animal = infodata[i][detailcols.ANIMAL];
                if (animal === "Yes") {
                    result = true;
                }
            }
            break;
        case desc.JESS_REMEMBER:
            {
                var animal = infodata[i][detailcols.JESS];
                if (animal === "Y") {
                    result = true;
                }
            }
            break;
        case desc.JESS_NO_REMEMBER:
            {
                var animal = infodata[i][detailcols.JESS];
                if (animal === "N") {
                    result = true;
                }
            }
            break;
        default:
            Logger.log("Invalid Descriptor: " + descriptor);
            break;
    }
  }

  return result;
} 

function populateHouseAvgStaticsByWeek(week)
{
    var staticsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HouseAvgStatic");
    var staticsheetdata = staticsheet.getDataRange().getValues();

    Logger.log("HouseAvg By Week " + week);

    var weekcol_0based = week + 1;
    // Range starts at 1,1 and is row,col
    var cell = staticsheet.getRange(1, weekcol_0based+1);
    cell.setValue(week);

    cell = staticsheet.getRange(2, weekcol_0based+1);
    cell.setValue(getAvgByDescriptorByWeek(1, week));
    cell = staticsheet.getRange(3, weekcol_0based+1);
    cell.setValue(getAvgByDescriptorByWeek(2, week));
    cell = staticsheet.getRange(4, weekcol_0based+1);
    cell.setValue(getAvgByDescriptorByWeek(3, week));
    
}

function populateHomerStaticsByWeek(week)
{
    var staticsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HomerStatic");
    var staticsheetdata = staticsheet.getDataRange().getValues();

    var weekcol_0based = week + 1;
    // Range starts at 1,1 and is row,col
    var cell = staticsheet.getRange(1, weekcol_0based+1);
    cell.setValue(week);

    cell = staticsheet.getRange(2, weekcol_0based+1);
    cell.setValue(homeFieldAdvantageByWeek(5, week));
    cell = staticsheet.getRange(3, weekcol_0based+1);
    cell.setValue(homeFieldAdvantageByWeek(6, week));
    cell = staticsheet.getRange(4, weekcol_0based+1);
    cell.setValue(homeFieldAdvantageByWeek(7, week));
    cell = staticsheet.getRange(5, weekcol_0based+1);
    cell.setValue(homeFieldAdvantageByWeek(8, week));
    cell = staticsheet.getRange(6, weekcol_0based+1);
    cell.setValue(homeFieldAdvantageByWeek(9, week));
    cell = staticsheet.getRange(7, weekcol_0based+1);
    cell.setValue(homeFieldAdvantageByWeek(10, week));
    
}

function populateDeltaStaticsByWeek(week)
{
    var staticsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DeltaStatic");
    var staticsheetdata = staticsheet.getDataRange().getValues();

    var weekcol_0based = week + 1;
    // Range starts at 1,1 and is row,col
    var cell = staticsheet.getRange(1, weekcol_0based+1);
    cell.setValue(week);

    cell = staticsheet.getRange(2, weekcol_0based+1);
    cell.setValue(getVotingDiffsByWeek(8,7, week));
    cell = staticsheet.getRange(3, weekcol_0based+1);
    cell.setValue(getVotingDiffsByWeek(8,5, week));
    cell = staticsheet.getRange(4, weekcol_0based+1);
    cell.setValue(getVotingDiffsByWeek(10,6, week));
    cell = staticsheet.getRange(5, weekcol_0based+1);
    cell.setValue(getVotingDiffsByWeek(9,6, week));
    cell = staticsheet.getRange(6, weekcol_0based+1);
    cell.setValue(getVotingDiffsByWeek(9,8, week));
    cell = staticsheet.getRange(7, weekcol_0based+1);
    cell.setValue(getVotingDiffsByWeek(10,8, week));
    cell = staticsheet.getRange(8, weekcol_0based+1);
    cell.setValue(getVotingDiffsByWeek(7,5, week));
    cell = staticsheet.getRange(9, weekcol_0based+1);
    cell.setValue(getVotingDiffsByWeek(8,6, week));
    cell = staticsheet.getRange(10, weekcol_0based+1);
    cell.setValue(getVotingDiffsByWeek(10,9, week));
    cell = staticsheet.getRange(11, weekcol_0based+1);
    cell.setValue(getVotingDiffsByWeek(6,5, week));
    cell = staticsheet.getRange(12, weekcol_0based+1);
    cell.setValue(getVotingDiffsByWeek(10,7, week));
    cell = staticsheet.getRange(13, weekcol_0based+1);
    cell.setValue(getVotingDiffsByWeek(10,5, week));
    cell = staticsheet.getRange(14, weekcol_0based+1);
    cell.setValue(getVotingDiffsByWeek(9,7, week));
    cell = staticsheet.getRange(15, weekcol_0based+1);
    cell.setValue(getVotingDiffsByWeek(9,5, week));
    cell = staticsheet.getRange(16, weekcol_0based+1);
    cell.setValue(getVotingDiffsByWeek(7,6, week));
}

function populateStatics()
{
  var latestweek = ratingdata.length - 1;
    
    var weektostart = latestweek;

    // Uncomment this to start with a specific week
    // Last was done after week 102
    //weektostart = 95;

    for(week = weektostart; week < ratingdata.length; week++ )
    {
      populateHomerStaticsByWeek(week);
      populateDeltaStaticsByWeek(week);
      populateHouseAvgStaticsByWeek(week);
    }
}

function getVotingDiffsByWeek(a, b, week)
{
    var num_weeks = ratingdata.length - 1; // Length has a heading row
    var weeksheet_string = "Week "+ String(week);
    if(week == num_weeks)
    {
      weeksheet_string = "Latest";
    }

    Logger.log("Getting voting diffs for week " + weeksheet_string);
    var weeksheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(weeksheet_string);
    var weeksheetdata = weeksheet.getDataRange().getValues();

    if(week < 58)
    {
      a = a - 3;
      b = b - 3;
    }

    var diff = 0;
    for (var i = 1; i < weeksheetdata.length; i++) {
        Logger.log("A: " + weeksheetdata[i][a] + ".  B: " + weeksheetdata[i][b]);
        diff += Math.abs(weeksheetdata[i][a] - weeksheetdata[i][b]);
    }

    Logger.log("Total Diff: " + diff);

    return diff;
}

function getVotingDiffs(a, b) {
  return getVotingDiffsByWeek(a,b, ratingdata.length-1);
}

function getDiffVsAge(a) {
    var diff = 0;
    for (var i = 1; i < ratingdata.length; i++) {
        diff += Math.abs(ratingdata[i][a] - ratingdata[i][Avg]);
    }

    return diff.toFixed(2);
}

function getAgreedMovie() {
    var agreed_movie_num = 0;
    var agreed_movie_diff = 10000;
    var total_diff = 0;

    Logger.log("In Get Agreed Movie");

    for (var i = 1; i < ratingdata.length; i++) {
        total_diff = 0;
        for (var j = Scott; j <= Pat; j++) {
            total_diff += Math.abs(ratingdata[i][Avg] - ratingdata[i][j]);
        }
        Logger.log("Total Diff for " + ratingdata[i][MovieName] + " is " + total_diff);
        if (total_diff < agreed_movie_diff) {
            agreed_movie_num = i;
            agreed_movie_diff = total_diff;
        }
    }

    return ratingdata[agreed_movie_num][MovieName];
}

function brockHome() {

  var num = 0;
  num =RtAvgByDescriptor(17);
  Logger.log("Brock Home is " + num);
  return;
}
function getControversialMovie() {
    var agreed_movie_num = 0;
    var agreed_movie_diff = 0;
    var total_diff = 0;

    Logger.log("In Get Agreed Movie");

    for (var i = 1; i < ratingdata.length; i++) {
        total_diff = 0;
        for (var j = Scott; j <= Pat; j++) {
            total_diff += Math.abs(ratingdata[i][Avg] - ratingdata[i][j]);
        }
        Logger.log("Total Diff for " + ratingdata[i][MovieName] + " is " + total_diff);
        if (total_diff > agreed_movie_diff) {
            agreed_movie_num = i;
            agreed_movie_diff = total_diff;
        }
    }

    return ratingdata[agreed_movie_num][MovieName];
}

function getWeekByMovieName(movie){
  var week = 0;
    for (var i = 1; i < infodata.length; i++) {
        if (infodata[i][detailcols.MOVIENAME] == movie) {
            week = i;
            break;
        }
    }

    Logger.log("Found that movie " + movie + " is from week " + week);
    return week;
}

function isWeekAHomePick(week, a) {
    var ishome = false;

    Logger.log("Checking Movie Week " + week + " House (" + infodata[week][detailcols.HOUSE] + ")");

    if (((a == Brock) || (a == Lauren)) && (infodata[week][detailcols.HOUSE] == "Collins")) {
        Logger.log("It is a Collins");
        ishome = true;
    }
    else if (((a == Pat) || (a == Elaine)) && (infodata[week][detailcols.HOUSE] == "Schreck")) {
        Logger.log("It is a Schreck");
        ishome = true;
    }
    else if (((a == Scott) || (a == Jess)) && (infodata[week][detailcols.HOUSE] == "Risolute")) {
        Logger.log("It is a Risolute");
        ishome = true;
    }

    return ishome;
}

function CollinsAverage() {
    return HomePickAvg(4);
}

function AvgAvg() {
    var total = 0;
    var count = 0;
    var average = 0;
    for (var i = 1; i < ratingdata.length; i++) {
        count++;
        Logger.log("Before Movie: " + ratingdata[i][MovieName] + " Total was " + total);
        total += ratingdata[i][Avg];
        Logger.log("After Movie: " + ratingdata[i][MovieName] + " TOtal was " + total);
    }

    average = total / count;
    Logger.log("Total:" + total + " Count:" + count + " Avg:" + average);

    return (average.toFixed(2));
}

function testHouseAverage() {
  var i;
  i=getAvgByDescriptor(1);
  i=getAvgByDescriptor(2);
  i=getAvgByDescriptor(3);
}
function getAvgByDescriptor(a) {
  return getAvgByDescriptorByWeek(a, ratingdata.length-1);
}

function getAvgByDescriptorByWeek(a, week) {
    var list = add_movies_by_desc(a, week);
    var total = 0;
    for (var i = 0; i < list.length; i++) {
        total += infodata[list[i]][detailcols.AVGRATING + week - 1];
        Logger.log("Adding Movie : " + list[i] + " Score: " + infodata[list[i]][detailcols.AVGRATING + week - 1] + " Total is now: " + total);

    }

    Logger.log("List Length is " + list.length + ". Total is : " + total);
    var toreturn = 0;

    if(list.length > 0)
    {
      var avg = total / (list.length);
      toreturn = avg.toFixed(2);
      Logger.log("Avg " + avg + " : ToRet: " + toreturn);
    }

    return toreturn;
}

function ImdbAvgTest() {
    return RtAvgByDescriptor(desc.HORROR);
}

function ImdbAvgByDescriptor(a) {
    var list = add_movies_by_desc(a, ratingdata.length-1);
    var total = 0;
    for (var i = 0; i < list.length; i++) {
        Logger.log("Adding Movie : " + list[i] + " with imdb: " + infodata[list[i]][detailcols.IMDB]);
        total += infodata[list[i]][detailcols.IMDB];
    }

    Logger.log("List Length is " + list.length + ". Total is : " + total);
    var avg = total / (list.length);
    return avg.toFixed(2);
}

function RtAvgByDescriptor(a) {
    var list = add_movies_by_desc(a, ratingdata.length - 1);
    var total = 0;
    var count = 0;
    var score = 0;
    for (var i = 0; i < list.length; i++) {
        Logger.log("Adding Movie : " + list[i] + "with rt: " + infodata[list[i]][detailcols.RT]);
        if (infodata[list[i]][detailcols.RT] !== "n/a") {
            total += (infodata[list[i]][detailcols.RT]) * 100;
            count++;
        }
    }

    Logger.log("List Length is " + list.length + ". Total is : " + total);
    var avg = 0;
    if (count !== 0) {
        avg = (total / count);
    }
    return avg.toFixed(2) + "%";
}

function HomePickAvg(a) {
    var total = 0;
    var count = 0;
    for (var i = 1; i < ratingdata.length; i++) {
        if (true == isWeekAHomePick(ratingdata[i][RatingWeekCol], a)) {
            count++;
            Logger.log("Before Movie: " + ratingdata[i][MovieName] + " Total was " + total);
            total += ratingdata[i][Avg];
            Logger.log("After Movie: " + ratingdata[i][MovieName] + " TOtal was " + total);
        }
    }
    Logger.log("Total:" + total + " Count:" + count);
    return (total / count).toFixed(3);
}

function testMovieDelta()
{
  Logger.log(biggestMovieDelta(5,true));
  return;
}
function biggestMovieDelta(person, high_flag)
{
  var winning_movie = "Error?";
  var winning_number = 0;
  var diff = 0;
  for (var i = 1; i < ratingdata.length; i++) {

    if(high_flag == true)
    {
        if(ratingdata[i][Avg] > ratingdata[i][person])
        {
          diff = Math.abs(ratingdata[i][person] - ratingdata[i][Avg]);
  Logger.log("Before Movie: " + ratingdata[i][MovieName] + " at " + diff + " Winner was " + winning_movie + " at " + winning_number);
          if(diff > winning_number)
          {
              winning_number = diff;
              winning_movie = ratingdata[i][MovieName] + " (" + winning_number.toFixed(2) + ")";
          }
        }
    }
    else
    {
        if(ratingdata[i][Avg] < ratingdata[i][person])
        {
          diff = Math.abs(ratingdata[i][person] - ratingdata[i][Avg]);
          if(diff > winning_number)
          {
              winning_number = diff;
              winning_movie = ratingdata[i][MovieName] + " (" + winning_number.toFixed(2) + ")";

          }
        }
    }
  }
  return winning_movie;
}

function getDescAvgForPerson(a, desc) {
    var total = 0;
    var count = 0;

    for (var i = 1; i < ratingdata.length; i++) {
        if (true == isWeekValidForDesc(ratingdata[i][RatingWeekCol], desc)) {
            count++;
            Logger.log("Before Movie: " + ratingdata[i][MovieName] + " Total was " + total);
            total += ratingdata[i][a];
            Logger.log("After Movie: " + ratingdata[i][MovieName] + " TOtal was " + total);
        }
    }
    Logger.log("Total:" + total + " Count:" + count);
    return (total / count).toFixed(3);
}

function homeFieldAdvantage(a) {
    var diff = homeFieldAdvantageByWeek(a, ratingdata.length - 1);
    return diff;
}

function GetCountOfTwoDescs(d1, d2) {
    var count = 0;

    for (var i = 1; i < infodata.length; i++) {
        if ((true == DoesMovieDescMovieCountForDesc(i, d1)) &&
            (true == DoesMovieDescMovieCountForDesc(i, d2))) {
            count++;
        }
    }

    Logger.log("Final Count is " + count);
    return count;
}

function DoesMovieDescMovieCountForDesc(i, descriptor) {
    var does_apply = false;
    if (infodata[i][detailcols.MOVIENAME] !== "") {
        switch (descriptor) {
            case desc.HOUSE_COLLINS:
                {
                    if (infodata[i][detailcols.HOUSE] == "Collins") {
                        does_apply = true;
                    }
                }
                break;
            case desc.HOUSE_RISOLUTE:
                {
                    if (infodata[i][detailcols.HOUSE] == "Risolute") {
                        does_apply = true;
                    }
                }
                break;
            case desc.HOUSE_SCHRECK:
                {
                    if (infodata[i][detailcols.HOUSE] == "Schreck") {
                        does_apply = true;
                    }
                }
                break;
            case desc.KUNGFU:
                {
                    Logger.log("Checking KungFu. " + infodata[i][detailcols.KUNGFU]);
                    if (infodata[i][detailcols.KUNGFU] == "Yes") {
                        does_apply = true;
                    }
                }
                break;
            case desc.ROMCOM:
                {
                    if (infodata[i][detailcols.ROMCOM] == "Yes") {
                        does_apply = true;
                    }
                }
                break;
            case desc.HORROR:
                {
                    if (infodata[i][detailcols.HORROR] == "Yes") {
                        does_apply = true;
                    }
                }
                break;
            case desc.XMAS:
                {
                    if (infodata[i][detailcols.XMAS] == "Yes") {
                        does_apply = true;
                    }
                }
                break;
            case desc.SCIFI:
                {
                    if (infodata[i][detailcols.SCIFI] == "Yes") {
                        does_apply = true;
                    }
                }
                break;
            case desc.ACTION:
                {
                    if (infodata[i][detailcols.ACTION] == "Yes") {
                        does_apply = true;
                    }
                }
                break;
            case desc.BOOBS:
                {
                    if (infodata[i][detailcols.BOOBS] == "Yes") {
                        does_apply = true;
                    }
                }
                break;
            case desc.NOBOOBS:
                {
                    if (infodata[i][detailcols.BOOBS] == "No") {
                        does_apply = true;
                    }
                }
                break;
            case desc.SEQUEL:
                {
                    if (infodata[i][detailcols.SEQUEL] == "Yes") {
                        does_apply = true;
                    }
                }
                break;
            case desc.SEQUEL:
                {
                    if (infodata[i][detailcols.SEQUEL] == "Yes") {
                        does_apply = true;
                    }
                }
                break;
            case desc.NOT_RATED:
                {
                    if (infodata[i][detailcols.RATING] == "Not Rated") {
                        does_apply = true;
                    }
                    if (infodata[i][detailcols.RATING] == "Unrated") {
                        does_apply = true;
                    }
                    if (infodata[i][detailcols.RATING] == "N/A") {
                        does_apply = true;
                    }
                }
                break;
            case desc.RATED_PG:
                {
                    if (infodata[i][detailcols.RATING] == "PG") {
                        does_apply = true;
                    }
                }
                break;
            case desc.RATED_PG13:
                {
                    if (infodata[i][detailcols.RATING] == "PG-13") {
                        does_apply = true;
                    }
                }
                break;
            case desc.RATED_TV14:
                {
                    if (infodata[i][detailcols.RATING] == "TV-14") {
                        does_apply = true;
                    }
                }
                break;
            case desc.RATED_R:
                {
                    if (infodata[i][detailcols.RATING] == "R") {
                        does_apply = true;
                    }
                }
                break;
            case desc.R_1980s:
                {
                    var year = infodata[i][detailcols.YEAR];
                    if ((year > 1979) && (year < 1990)) {
                        does_apply = true;
                    }
                }
                break;
            case desc.R_1990s:
                {
                    var year = infodata[i][detailcols.YEAR];
                    if ((year > 1989) && (year < 2000)) {
                        does_apply = true;
                    }
                }
                break;
            case desc.R_2000s:
                {
                    var year = infodata[i][detailcols.YEAR];
                    if ((year > 1999) && (year < 2010)) {
                        does_apply = true;
                    }
                }
                break;
            case desc.R_2010s:
                {
                    var year = infodata[i][detailcols.YEAR];
                    if ((year > 2009) && (year < 2020)) {
                        does_apply = true;
                    }
                }
                break;
            case desc.LESS_90:
                {
                    var time = infodata[i][detailcols.RUNTIME];
                    var array1 = [{}];

                    array1 = time.split(" ");
                    Logger.log(infodata[i][detailcols.MOVIENAME] + " runtime " + array1[0]);
                    if (array1[0] < 91) {
                        does_apply = true;
                    }
                }
                break;
            case desc.MORE_90:
                {
                    var time = infodata[i][detailcols.RUNTIME];
                    var array1 = [{}];

                    array1 = time.split(" ");
                    Logger.log(infodata[i][detailcols.MOVIENAME] + " runtime " + array1[0]);
                    if (array1[0] > 90) {
                        does_apply = true;
                    }
                }
                break;
            case desc.ANIMAL_ATTACK:
                {
                    var animal = infodata[i][detailcols.ANIMAL];
                    if (animal === "Yes") {
                        does_apply = true;
                    }
                }
                break;
            case desc.JESS_REMEMBER:
                {
                    var animal = infodata[i][detailcols.JESS];
                    if (animal === "Y") {
                        does_apply = true;
                    }
                }
                break;
            case desc.JESS_NO_REMEMBER:
                {
                    var animal = infodata[i][detailcols.JESS];
                    if (animal === "N") {
                        does_apply = true;
                    }
                }
                break;
            default:
                does_apply = false;
                break;
        }

        Logger.log("Movie " + infodata[i][detailcols.MOVIENAME] + " Result:" + does_apply);
    }
    return does_apply;
}

function avgBoobRating() {
    var total = 0;
    var count = 0
    for (var i = 1; i < infodata.length; i++) {
        if (infodata[week][detailcols.BOOBS] == "Yes") {
            total += infodata[i][detailcols.AVGRATING + ratingdata.length - 1 - 1];
            count++;
        }

    }
    Logger.log("Boob Count: " + count);

    return (total / count).toFixed(3);
}

function avgNonBoobRating() {
    var total = 0;
    var count = 0
    for (var i = 1; i < infodata.length; i++) {
        if (infodata[week][detailcols.BOOBS] == "No") {
            total += infodata[i][detailcols.AVGRATING + ratingdata.length - 1 - 1];
            count++;
        }

    }
    Logger.log("Nonboob Count: " + count);
    return (total / count).toFixed(3);
}
