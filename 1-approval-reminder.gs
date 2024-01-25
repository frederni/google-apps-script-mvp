function send_approval_reminder(){
  const demonstration = true;
  const dueDate = 5;
  const status = 2;
  let recipient = {
    "Til godkjenning": "cfo@soprasteria.com",
    "Til nettbank": "accounting@soprasteria.com"
  };
  // Don't mind the code below -- just for demonstration purposes
  Object.keys(recipient).forEach(key => {recipient[key] = "frederick.nilsen+"+recipient[key];});

  // Initialization / Load values
  let ss = SpreadsheetApp.getActiveSheet();
  let today = new Date();
  // DEBUG
  today = new Date("2024-02-03");
  let needsAttention = {"Til godkjenning": [], "Til nettbank": []};
  let data = ss.getRange("A4:F1000").getDisplayValues();

  // Go through each row
  for (let i=0; i < data.length; i++){
    let dueDateRow = new Date(Date.parse(
      data[i][dueDate].split('.').reverse().join('-')
      ));
    let timeDifference = (dueDateRow - today) / (1000.0 * 60.0 * 60.0 * 24.0);
    // Check if row requires email reminder
    if (data[i][status] !== 'Betalt' && timeDifference < 3.0){
      needsAttention[data[i][status]].push(data[i]);
    }
  }
  console.log(needsAttention);
  
  // Construct message
  for (const actionType in needsAttention){
    let message = "Hei!<br />Dette er en automatisk påminnelse på at du har betalinger " + actionType.toLowerCase() + ":<br /><br />";
    message += "<table><tr><th>FID</th> <th>Leverandør</th> <th>Status</th> <th>Sum</th> <th>Fakturadato</th> <th>Forfallsdato</th>";
    for( let i = 0; i < needsAttention[actionType].length; i++){
      let row = "<tr>"
      for (let j = 0; j < needsAttention[actionType][i].length; j++){
        row += "<td>" + needsAttention[actionType][i][j] + "</td>";
      }
      message += row + "</tr>";
    }
    message += "Skrevet av NetPulse bot! 🤖💙";
    if(needsAttention[actionType].length > 0){
      MailApp.sendEmail({
        to: recipient[actionType],
        subject: "Automatisk påminnelse: Fakturaer " + actionType.toLowerCase(),
        htmlBody: message
      });
      console.log("Email sent to " + recipient[actionType]);
    }

  }

}
