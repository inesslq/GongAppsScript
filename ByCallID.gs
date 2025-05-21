// 📄 Gong Transcript Inserter for Google Docs
// This Google Apps Script allows users to fetch a Gong call transcript by entering a Call ID, and automatically insert it into a Google Doc. It includes call metadata (date, URL) and correctly maps speaker names to their statements.

const API_KEY = "INSERT YOUR API ACCESS KEY HERE"; // ‼️ Insert your GONG API key here
const API_SECRET = "INSERT YOUR SECRET HERE"; // ‼️ Insert your GONG API secret here

// Functions to add a custom menu, prompting a user to insert a call id for the transcript lookup

function onOpen() {
  var ui = DocumentApp.getUi();
  ui.createMenu("Gong API")
    .addItem("Insert Call ID", "showPrompt")  // Adds a single menu item to the Google Docs UI
    .addToUi();
}

function showPrompt() {
  var ui = DocumentApp.getUi();
  var response = ui.prompt("Enter Call ID", "Please enter the Gong Call ID:", ui.ButtonSet.OK_CANCEL); // Menu items prompts a user to enter a Gong Call ID

  if (response.getSelectedButton() == ui.Button.OK) {
    var callId = response.getResponseText().trim();
    if (callId) {
      fetchAndAppendTranscripts(callId);
    } else {
      ui.alert("Invalid input. Please enter a Call ID.");
    }
  }
}

// Function to pull transcript data along with matched speaker names (instead of just speaker IDs)

function fetchAndAppendTranscripts(callId) {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();

  var apiKey = API_KEY; 
  var apiSecret = API_SECRET; 
  var authHeader = "Basic " + Utilities.base64Encode(apiKey + ":" + apiSecret);

  // Fetch Speaker Details (parties) using exposedFields
  var speakerUrl = "https://us-13059.api.gong.io/v2/calls/extensive";
  var speakerPayload = JSON.stringify({
    "filter": { "callIds": [callId] },
    "contentSelector": {
  "exposedFields": {
    "parties": true
  }
} 
  });

  var speakerOptions = {
    method: "post",
    headers: {
      "Authorization": authHeader,
      "Content-Type": "application/json"
    },
    payload: speakerPayload,
    muteHttpExceptions: true
  };

  var speakerResponse = UrlFetchApp.fetch(speakerUrl, speakerOptions);
  var speakerJson = JSON.parse(speakerResponse.getContentText());

  Logger.log("🔍 Speaker API Response: " + JSON.stringify(speakerJson, null, 2));

  if (!speakerJson.calls || speakerJson.calls.length === 0) {
    Logger.log("⚠️ No call details found.");
    return;
  }

  var speakerMap = {};
  var parties = speakerJson.calls[0]?.parties || [];

  if (parties.length === 0) {
    Logger.log("⚠️ No parties found in API response.");
  }

  parties.forEach(function (party) {
    if (party.speakerId && party.name) {
      speakerMap[party.speakerId] = party.name;
    }
  });

  Logger.log("🎙️ Speaker Mapping: " + JSON.stringify(speakerMap, null, 2));

  // Fetch Transcript
  var transcriptUrl = "https://us-13059.api.gong.io/v2/calls/transcript";
  var transcriptPayload = JSON.stringify({ "filter": { "callIds": [callId] } });

  var transcriptOptions = {
    method: "post",
    headers: {
      "Authorization": authHeader,
      "Content-Type": "application/json"
    },
    payload: transcriptPayload,
    muteHttpExceptions: true
  };

  var transcriptResponse = UrlFetchApp.fetch(transcriptUrl, transcriptOptions);
  var transcriptJson = JSON.parse(transcriptResponse.getContentText());

  Logger.log("🔍 Transcript API Response: " + JSON.stringify(transcriptJson, null, 2));

  if (!transcriptJson.callTranscripts || transcriptJson.callTranscripts.length === 0) {
    Logger.log("⚠️ No call transcripts found.");
    return;
  }

  var transcriptData = transcriptJson.callTranscripts[0]?.transcript || [];

  if (transcriptData.length === 0) {
    Logger.log("⚠️ No transcript data found.");
    return;
  }

  var rawDate = new Date(speakerJson.calls[0]?.metaData.scheduled);
  // Format the date to "Month Day, Year"
  var options = { year: 'numeric', month: 'long', day: 'numeric' };
  var formattedDate = rawDate.toLocaleDateString('en-US', options);
  var callURL = speakerJson.calls[0]?.metaData.url;

  var transcriptText ="\n________________NEW CALL______________\n\n" + "\n📞 Call ID: " + callId + "\n📆 Call Date: " + formattedDate + "\n🔗 Call URL: " + callURL +"\n\n";

  transcriptData.forEach(function (part) {
    var speakerName = speakerMap[part.speakerId] || "Unknown Speaker";
    
    part.sentences.forEach(function (sentence) {
      transcriptText += speakerName + ": " + sentence.text + "\n";
    });
  });

  if (transcriptText.trim() !== "") {
    body.appendParagraph(transcriptText);
    Logger.log("✅ Transcript added to Google Doc!");
  } else {
    Logger.log("⚠️ No transcript text found.");
  }
}
