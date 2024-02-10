/* WARNING: This script requires the Fuse.js library. For this to work, you copy the content of:
https://cdn.jsdelivr.net/npm/fuse.js/dist/fuse.js
into a new .gs-file in your Apps Script library. Ensure that the new file is above the main file.
The OpenAI API key must obviosuly be configured too. 
*/

function get_gpt_payload(prompt) {

  var modelID = "gpt-3.5-turbo";
  var temperature = 0.7;
  // Build the API payload
  return {
    'model': modelID,
    'temperature': temperature,
    'messages': [{"role": "user", "content": prompt}]
  };
}

function get_gpt_response(payload, API_KEY) {
  var options = {
      "method": "post",
      "headers": {
        "Content-Type": "application/json",
        "Authorization": "Bearer " + API_KEY
      },
      "payload": JSON.stringify(payload)
    };
  var response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", options);
  return JSON.parse(response.getContentText()).choices[0].message.content.trim();
}



function isSimilar(str1, str2) {
  if (str1 === str2) {
    return true;
  }
  const fuse = new Fuse([str1], { threshold: 0.2, distance: 10 });
  let ret = fuse.search(str2).length > 0;
  if(ret){
    Logger.log("##isSimilar instance returned true for '" + str1 + "' and '" + str2 + "'.");
  }
  return ret;
}

function bronzeFetchAllDinners() {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("dinners").clear();
  let folders = DriveApp.getFolderById("<FOLDERID>").getFolders();
  let offset = 0;
  while (folders.hasNext()){
    let filesInFolder = folders.next().getFiles();
    while (filesInFolder.hasNext()){
      let currentSpreadsheet = filesInFolder.next();
      Logger.log("Henter middager fra: " + currentSpreadsheet.getName());
      let dinners = SpreadsheetApp.open(currentSpreadsheet).getActiveSheet().getRange("C2:F100").getValues();
      let numCols = 1 + "F".charCodeAt(0) - "C".charCodeAt(0);
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("dinners").getRange(1+offset,1,dinners.length,numCols).setValues(dinners);
      offset = offset + dinners.length;
    }
  }
}

function silverCleanDinners() {
  const idxDate = 0;
  const idxType = 1;
  const idxDinner1 = 2;
  const idxDinner2 = 3;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("dinners");
  var range = sheet.getDataRange();
  var values = range.getValues();
  Logger.log("Removing blank values. Length before: " + values.length);
  values = values.filter(row => row.every(cell => cell !== "")) // All columns in row are blank
                .filter(row => row[idxDate] !== "") // No date
                .filter(row => row[idxType] !== "") // No dinnerType (meat, fish etc)
                .filter(row => row[idxDinner1] !== "" && row[idxDinner2] !== ""); // No dinner1 and no dinner2
  Logger.log("Values after removing blanks: " + values.length);

  // Unpivot dinner1 and dinner2
  Logger.log("Unpivoting data...")
  var newData = [];
  values.forEach(row => {
    newData.push([row[idxDate], row[idxType], row[idxDinner1]]);
    if (row[idxDinner2] !== "") {
      newData.push([row[idxDate], row[idxType], row[idxDinner2]]);
    }
  });
  Logger.log("Length after unpivot: " + newData.length);
  const idxDinner = 2; // For clarity: just one dinner column after unpivot

  // Remove rows matching the excluded dinner list
  Logger.log("Removing dinners from exclude-list");
  let excludeItems = SpreadsheetApp.getActiveSpreadsheet()
                      .getSheetByName("explicit_ignore")
                      .getRange("A:A")
                      .getValues()
                      .flat()
                      .filter(cell => cell !== "");

  Logger.log("Expected to remove " + excludeItems.length + " items (length of exclude list)");
  newData = newData.filter(row => excludeItems.indexOf(row[idxDinner]) === -1);
  Logger.log("Length after exclude list is " + newData.length);
  
  // Strip dinner values
  Logger.log("Stripping whitespaces for dinner names");
  newData.forEach(row => {
    row[idxDinner] = row[idxDinner].trim();
  });

  // Duplicate removal
  Logger.log("Duplicate removal");
  newData = newData.filter((row, index, self) => {
      return (
        index === self.findIndex(
          r => r[idxType] === row[idxType] && isSimilar(r[idxDinner], row[idxDinner])
        )
      );
    });
  
  // Clear existing data and paste the cleaned data
  sheet.clear();
  sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
}

function dataPipeline(){
  bronzeFetchAllDinners();
  silverCleanDinners();
}

function getSuggestion(){
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let mainSheet = ss.getActiveSheet();
  let dinnerType = mainSheet.getRange("L28").getValue();
  let dinnerSeason = mainSheet.getRange("L29").getValue();
  let filteredValues = ss.getSheetByName("dinners").getRange("A1:C").getDisplayValues();
  if (dinnerSeason === "Sommer"){ // April through september
    filteredValues = filteredValues.filter(
      function(row){
      let month = parseInt(row[0].split(".")[1], 10);
      return row[1] === dinnerType && month >= 4 && month <= 9;
      }
    )
  }
  else if(dinnerSeason === "Vinter"){ // October through March
    filteredValues = filteredValues.filter(
      function(row){
      let month = parseInt(row[0].split(".")[1], 10);
      return row[1] === dinnerType && month >= 4 && month <= 9;
      }
    )
  }
  else{ // Don't filter on season
    filteredValues = filteredValues.filter(
      function(row){
      let month = parseInt(row[0].split(".")[1], 10);
      return row[1] === dinnerType && month >= 4 && month <= 9;
      }
    )
  }
  // Choose randomly from potential dishes
  let randomIndex = Math.floor(Math.random() * filteredValues.length);
  Logger.log(filteredValues[randomIndex][1]);
  ss.getActiveSheet().getRange("N29").setValue(filteredValues[randomIndex][2]);
}


function getChatGPTsuggestion(){
  const DEBUG = false;
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getActiveSheet();
  let dinnerType = sheet.getRange("L37").getValue();
  let promptAppend;
  console.log(dinnerType);
  switch (dinnerType){
    case "Kjøtt":
      promptAppend = "- The recipe must contain meat, preferably white meat and not fish.";
      break;
    case "Fisk":
      promptAppend = "- The recipe must contain fish or seafood.";
      break;
    case "Vegetar":
      promptAppend = "- The recipe must be vegetarian.";
      break;
    default:
      promptAppend = "";
  }
  let gptPrompt = `Suggest a dinner recipe. The format should start with the name of the dish, followed by newline, followed by the ingredients and instructions. The suggestion must abide the following rules:
- The recipe is indended for 2-3 people
- The ingredients should be easily available in Norway
- The recipe should not take more than 30 minutes to prepare
- All measurement units should be in SI units
` + promptAppend;
  if(sheet.getRange("K39").getValue() !== ""){
    gptPrompt = gptPrompt + "\n- " + sheet.getRange("K39").getValue();
  }

  Logger.log("Submitting prompt to API: \n" + gptPrompt);
  let output;
  if (! DEBUG){
    output = get_gpt_response(get_gpt_payload(gptPrompt), "<APIKEY>");
  }
  else{
    output = `Lorem ipsum dolor sit amet, consectetur adipiscing elit. Vivamus risus mi, sagittis a lorem ac, convallis ullamcorper felis.
    Phasellus sodales purus lectus, at commodo lectus aliquam eget. Ut sit amet eleifend felis. Curabitur dignissim ligula lacus, eget vehicula eros scelerisque ut. Vivamus accumsan, velit vitae dignissim tincidunt, nulla nulla volutpat purus, quis scelerisque turpis libero ac ante. Quisque augue diam, molestie quis tincidunt vel, egestas eu dui.
    Nam non orci id mauris interdum tincidunt. Suspendisse efficitur faucibus nisl. Phasellus semper eros turpis. Mauris tempor justo ut nibh volutpat, ac ultrices lacus pretium. Nulla massa ante, faucibus sit amet tristique eget, placerat non nunc. Morbi dignissim dui ac orci iaculis egestas.`;
  }
  Logger.log(output);
  
  // Display output
  let arrayOutput = output.split("\n");
  sheet.getRange("O36").setValue(arrayOutput[0]);
  sheet.getRange("O37").setValue(arrayOutput.slice(1).join("\n"));
}

function saveGPTsuggestion(){
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getActiveSheet();
  let dishName = sheet.getRange("O36").getValue();
  let dishDescription = sheet.getRange("O37").getValue().trim();
  let dinnerType = sheet.getRange("L37").getValue();
  let now = new Date();
  let recipeSheet = ss.getSheetByName("Lagrede oppskrifter");
  recipeSheet.appendRow([now, dinnerType, dishName, dishDescription]);

  let savedRow = recipeSheet.getRange("C1:C").getValues().filter(String).length+1;
  let ssId = ss.getId().toString();
  let sheetId = recipeSheet.getSheetId().toString();
  sheet.getRange("O36").setValue("Recipe saved to:");
  sheet.getRange("O37").setFormula("=HYPERLINK(\"https://docs.google.com/spreadsheets/d/" + ssId + "/edit#gid=" + sheetId + "&range=C" + savedRow.toString() + "\"; \"" + dishName + "\")");
  sheet.getRange("O37").setFontColor("blue");
  sheet.getRange("O37").setFontLine('underline');
}
