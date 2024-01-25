function buildReport() {
  const ss = SpreadsheetApp.getActive();
  let data = ss.getActiveSheet().getRange("A4:D100").getValues();
  let payloadList = buildAlert(data);
  sendAlert(payloadList);
}

function nameToTag(name){
  
  let map_nameID = { 
    "SlackName": "U0123456789"
  }
  tag = map_nameID[name];

  if(tag){return "<@"+tag+">";}
  return name;
}

function buildAlert(data) {
  let today = new Date();
  let name = "Ukjent";
  let additional_info = "";
  let meetingToday = false;
  for(let i=0; i<data.length;i++){
    let rowDate = new Date(data[i][1]);
    if(today.getFullYear() == rowDate.getFullYear() & today.getMonth() == rowDate.getMonth() & today.getDate() == rowDate.getDate()){
      name = data[i][2];
      additional_info = data[i][3];
      if(name.length==0 && extraName.length==0){
        meetingToday = false;
      }
      else{
        meetingToday = true;
      }
    }
  }

  name = nameToTag(name);
  if(additional_info != ""){additional_info = "Kommentar: " + additional_info;}

  let payload = {
    "blocks": [
      {
        "type": "section",
        "text": {
          "type": "mrkdwn",
          "text": ":cake: *Snacks til møte i dag i dag* :cake:"
        }
      },
      {
        "type": "divider"
      },
      {
        "type": "section",
        "text": {
          "type": "mrkdwn",
          "text": name + " er ansvarlig for møtesnacks i dag. " + additional_info
        }
      },
      {
        "type": "section",
        "text": {
          "type": "mrkdwn",
          "text": "*Sendt av Møtesnack-bot* :robot_face: "
        }
      }
    ]
  };
  return [payload, meetingToday];
}

function sendAlert(payloadList) {
  const webhook ="https://hooks.slack.com/services/T00000000/B00000000/XXXXXXXXXXXXXXXXXXXXXXXX";
  var options = {
    "method": "post", 
    "contentType": "application/json", 
    "muteHttpExceptions": true, 
    "payload": JSON.stringify(payloadList[0]) 
  };
  if(payloadList[1]){
    try {
      UrlFetchApp.fetch(webhook, options);
    } catch(e) {
      Logger.log(e);
    }
  }
}
