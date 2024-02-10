# Google Apps Script MVP
*A showcase of ~3~ 4 use cases for Google Apps Script*

Many people don't know how powerful Google Apps Script really is. Considering it's free for end-users with a gmail account, it can boost your productivity and automate processes in Google Drive, Sheets, Forms and other Google Workspace services. Below are 3 specific use cases that are easy to reimplement to your own needs.


## Usecase 1 - Email notification on surpassed deadline for invoice approval

This usecase assumes the following spreadsheet format:

![image](https://github.com/frederni/google-apps-script-mvp/assets/23258333/a1a96b08-8c9f-45d9-9fad-77812136cdba)

- Takes use of `MailApp` to send an email with HTML body based on cell values in spreadsheet.
- Segregates items requiring action by status, in turn sending separate mails to different recipient.
- Should be scheduled to run daily

## Usecase 2 - Slack webhook integration

This usecase assumes the following spreadsheet format:

![image](https://github.com/frederni/google-apps-script-mvp/assets/23258333/dd5dad08-9a62-4cc8-b69d-18d9cf619753)

- Uses `UrlFetchApp` and an incoming webhook from Slack to send a POST request with JSON defining a Slack message.
- To use, one must create a Slack workspace, a Slack app and define what channel the payload should be posted to.
- Uses the same "naive" logic of iterating through rows and checking if they match condition to send (date is equal to today)
- Should be scheduled to run daily

## Usecase 3 - Control panel

This usecase demonstrates the capabilities of `DriveApp`. In the example below, we use illustrations that can be assigned a script function, so that an end-user would only need to press a button for the script to trigger.

![image](https://github.com/frederni/google-apps-script-mvp/assets/23258333/ee9f2318-6a78-43bb-9b0c-e44729d86fef)

- "Select all" button and "Deselect all" button that just sets the defined range to true/false, thus updating checkboxes
- **Create budget** function that copies "template" sheet (whose Drive ID is xxx in example image) into folder (whose Drive ID is yyy)
- Option to overwrite existing sheet if Spreadsheet with matching ID is in destination folder
- Changes values from template after copy based on budget name
- **Change formula** function that takes in a range and formula and applied formula to each selected budget. `setFormula` will auto-increment formula unless `$` is part of range reference
- Should not have any scheduling

## Usecase 4 - Dinner suggestion and OpenAI integration

This is actually taken from a personal project I had. TL;DR I have a folder of spreadsheets with dinner dishes for each day. The usecase implements a data pipeline that:

- Fetches all dinners from previous spreadsheets (all in GDrive folder `<FOLDERID>`) in a bronze layer
- Cleans data, removing duplicates with fuzzy search, trims text, and removes exact matches from an exclude list
- Has a function to retrieve random recipe from dataset using parameters like dinner type ("Meat", "Fish", "Vegetarian") and season ("Summer", "Winter")
    - The season parameter will filter dinners that were had  
- OpenAI integration that prompts GPT 3.5 turbo to write dish and instructions
    - Option to save suggestion to other sheet, grabbing the cell URL to easily access
    - Prompt will change on dinner type (meat, fish, vegetarian) and can add customized message
    - Requires API key to work
