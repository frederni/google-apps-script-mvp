function myFunction() {
  let currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let currentSheet = current_spreadsheet.getActiveSheet();
  let myRange = current_sheet.getRange("A1");
  myRange.setValue("Hello world!");
}

function select_all(){
  let checkboxes = SpreadsheetApp.getActiveSheet().getRange("A4:A33");
  checkboxes.setValue("TRUE");
}
function deselect_all(){
  let checkboxes = SpreadsheetApp.getActiveSheet().getRange("A4:A33");
  checkboxes.setValue("FALSE");
}

function create_budget() {
  let ss = SpreadsheetApp.getActiveSheet();
  let budget_names_arr = ss.getRange("A4:B33").getValues();
  let destination_folder = DriveApp.getFolderById(ss.getRange("K5").getValue());
  let budget_template = DriveApp.getFileById(ss.getRange("K6").getValue());

  for (let i=0;i<budget_names_arr.length; i++){
    if(budget_names_arr[i][0]){ // Only for selected budgets
      let budget_name = budget_names_arr[i][1];
      let overwrite_file = ss.getRange("H7").getValue();

      if (overwrite_file){ // Deletes files with identical name
        let existing_files = destination_folder.getFilesByName(budget_name);
        while(existing_files.hasNext()){
          console.log("Budget", budget_name, "exists in folder and overwrite is true. Deleting...");
          existing_files.next().setTrashed(true);
        }
      }

      // Copy template
      let newfile = budget_template.makeCopy(budget_name, destination_folder);
      // Change name in cell of new file to match the budget name:
      SpreadsheetApp.open(newfile).getRange("C1").setValue(budget_name);
      console.log("Successfully made budget for department", budget_name);
    }
  }
}

function change_formula(){
  let ss = SpreadsheetApp.getActiveSheet();
  let formulaRange = ss.getRange("H12").getValue();
  let formula = ss.getRange("H13").getValue();
  let budget_names_arr = ss.getRange("A4:B33").getValues();

  for (let i = 0; i < budget_names_arr.length; i++){
    if(budget_names_arr[i][0]){
      let destination_folder = DriveApp.getFolderById(ss.getRange("K5").getValue());
      current_budget = destination_folder.getFilesByName(budget_names_arr[i][1]);
      while(current_budget.hasNext()){
        let budget_ss = SpreadsheetApp.open(current_budget.next()).getActiveSheet();
        budget_ss.getRange(formulaRange).setFormula("=" + formula);
        console.log("Changed formula for budget", budget_names_arr[i][1]);
      }
    }
  }

}
