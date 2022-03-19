//index from 1
var WEEK_COLUMN_NO = 1;
var SCORE_COLUMN_NO = 5;
var IMDB_SCORE_COLUMN_NO = 6;
var HOUSE_COLUMN_NO = 7;
var RATING_COLUMN_NO = 8;
var RUNTIME_COLUMN_NO = 9;
var BOXOFFICE_COLUMN_NO = 10;
var RELEASE_YEAR_COLUMN_NO = 11;
var BOOBS_COLUMN_NO = 12;
var KARATE_COLUMN_NO = 13;
var ACTION_COLUMN_NO = 14;
var HORROR_COLUMN_NO = 15;
var SCIFI_COLUMN_NO = 16;
var ROMCOM_COLUMN_NO = 17;
var XMAS_COLUMN_NO = 18;
var SEQUEL_COLUMN_NO = 20;
var ANIMAL_COLUMN_NO = 21;
var JESS_COLUMN_NO = 22;
var PLOT_COLUMN_NO = 4;
var TITLE_COLUMN_NO = 3;
var IMDB_ID_COLUMN_NO = 19;
var POSTER_COLUMN_NO = 2;
var WEEK_COLUMN_NO = 1;
var AVG_COLUMN_NO = 22;
var SHEET_NAME = "MovieDetails";
var RANK_SHEET = "Latest";

var LatestWeekCol = 11;
var LatestAvgCol = 4;

var API_URL = "https://www.omdbapi.com/?i=tt3896198&apikey=98240132&";
//var API_URL = "www.omdbapi.com/?apikey=[98240132]&";
var IMDB_URL = "www.imdb.com/title/";

//global
var ss = SpreadsheetApp.getActive();
var sh = ss.getSheetByName(SHEET_NAME);
var latestsheet = ss.getSheetByName(RANK_SHEET);
var lRow = sh.getLastRow(), lCol = sh.getLastColumn();

function   addMovieDetailRow(movie)
{
    var row = sh.getRange(movie.week + 1 , 1, 1, lCol);
    row.getCell(1, TITLE_COLUMN_NO).setValue(movie.title);
    row.getCell(1, IMDB_ID_COLUMN_NO).setValue(movie.imdb);
    row.getCell(1, HOUSE_COLUMN_NO).setValue(movie.house);
    row.getCell(1, BOOBS_COLUMN_NO).setValue(movie.boobs);
    row.getCell(1, KARATE_COLUMN_NO).setValue(movie.karate);
    row.getCell(1, ACTION_COLUMN_NO).setValue(movie.action);
    row.getCell(1, HORROR_COLUMN_NO).setValue(movie.horror);
    row.getCell(1, SCIFI_COLUMN_NO).setValue(movie.sci);
    row.getCell(1, ROMCOM_COLUMN_NO).setValue(movie.rom);
    row.getCell(1, XMAS_COLUMN_NO).setValue(movie.xmas);
    row.getCell(1, SEQUEL_COLUMN_NO).setValue(movie.seq);
    row.getCell(1, ANIMAL_COLUMN_NO).setValue(movie.animal);
    row.getCell(1, JESS_COLUMN_NO).setValue(movie.jess);
    row.getCell(1, WEEK_COLUMN_NO).setValue(movie.week);

  return;
}

function pullMovieDetailsIntoSpreadsheet(e){
  getEmptyScoreRows();
}

function getEmptyScoreRows(){  
  //Split into two lists so that empty scores are prioritized before checking n/a ones
  var rowNoList = [];
  var naList = [];
  
  var allScoresValues = sh.getRange(2, SCORE_COLUMN_NO, lRow, 1).getValues();
  for (var i = 0; i < allScoresValues.length; i++) {  
      rowNoList.push(i + 2);
  }

  addRTScores(rowNoList);  
}

function addAverages()
{
    var infosheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MovieDetails");
    var infodata = infosheet.getDataRange().getValues();
    var ratingsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Latest");
    var ratingdata = ratingsheet.getDataRange().getValues();
    var i = 0;
    var j = 0;
    var week = 0;
    var avg = 0;

    for(i = 1; i < infodata.length; i++)
    {
        week = infodata[i][WEEK_COLUMN_NO];
        for(j = 1; j < ratingdata.length; j++)
        {
            Logger.log("Checking movie " + week + " verses " + ratingdata[j][LatestWeekCol]);
            if(week == ratingdata[j][LatestWeekCol])
            {
                Logger.log("Found");
                avg = ratingdata[j][LatestAvgCol];

                var cell = infosheet.getRange(i+1,AVG_COLUMN_NO+ratingdata.length-1);
                cell.setValue(avg);
                //infodata[i][AVG_COLUMN_NO] = avg;
                break;
            }
        }
    }

    return;
}
function addRTScores(emptyScoreRowNos) {
  for (var i = 0; i < emptyScoreRowNos.length; i++){
    var row = sh.getRange(emptyScoreRowNos[i] , 1, 1, lCol);
    var titleCol = row.getCell(1, TITLE_COLUMN_NO);
    var plotCol = row.getCell(1, PLOT_COLUMN_NO);
    var imdbIdCol = row.getCell(1, IMDB_ID_COLUMN_NO);
    var ratingCol = row.getCell(1, RATING_COLUMN_NO);
    var runtimeCol = row.getCell(1, RUNTIME_COLUMN_NO);
    var boxofficeCol = row.getCell(1, BOXOFFICE_COLUMN_NO);
    var yearCol = row.getCell(1, RELEASE_YEAR_COLUMN_NO);
    var posterCol = row.getCell(1, POSTER_COLUMN_NO);
  
    var imageCol = row.getCell(1, POSTER_COLUMN_NO);
    Logger.log("AddRtScore:"+i);
      var params;

    if (titleCol.getValue() !== "" && posterCol.getValue() === ""){
            Logger.log("Title is "+titleCol.getValue()+". ImdCol is:"+imdbIdCol.isBlank());



      if(imdbIdCol.isBlank() === false)
      {
        params = getRTScoreById(imdbIdCol.getValue());
      }
      else
      {
        params = getRTScore(titleCol.getValue());
      }

      if(params.Response !== "False")
      {
        var score = "n/a";
      if(params.hasOwnProperty("Poster"))
      {
       imageCol.setFormula('IMAGE("'+params.Poster+'")');
       imageCol.setHorizontalAlignment("center");
      }
      else
      {
        Logger.log("No Poster");
      }
        if(params.Ratings.length > 1)
        {
       score = params.Ratings[1].Value;
        }
      
      var scoreCol = row.getCell(1, SCORE_COLUMN_NO);
      var imdbScoreCol = row.getCell(1, IMDB_SCORE_COLUMN_NO);
      if (score === undefined || score === "N/A"){
        scoreCol.setValue("N/A");
      }
      else
      {
        scoreCol.setValue(score);
      }

        imdbScoreCol.setValue(params.imdbRating);

        var plot = params.Plot;
        if (plot !== undefined)
        {
          plotCol.setValue(plot);
        }
        else
        {
          plotCol.setValue("Plot Unavailable");
        }

        ratingCol.setValue(params.Rated);

        boxofficeCol.setValue(params.BoxOffice);
        runtimeCol.setValue(params.Runtime);
        yearCol.setValue(params.Year);
      }
    }
  }
}

function getRTScore(title) {
  //Create URL
  var queryString = "?r=json&tomatoes=true&t=" + title;
  var url = API_URL + queryString;

  var options =
      {
        "method"  : "GET",   
        "followRedirects" : true,
        "muteHttpExceptions": true
      };
  
  //send GET request
  var result = UrlFetchApp.fetch(url, options);
  Logger.log("Result for URL:" + url+" is :"+result);

  if (result.getResponseCode() == 200) {
    return JSON.parse(result.getContentText());
  }
}

function getRTScoreById(id) {
  //Create URL
  var url = "https://www.omdbapi.com/?i="+id+"&apikey=98240132&r=json&tomatoes=true";

  var options =
      {
        "method"  : "GET",   
        "followRedirects" : true,
        "muteHttpExceptions": true
      };
  
  //send GET request
  var result = UrlFetchApp.fetch(url, options);
  Logger.log("Result for URL:" + url+" is :"+result);

  if (result.getResponseCode() == 200) {
    return JSON.parse(result.getContentText());
  }
}