// 📄 Gong Transcript Inserter for Google Docs
// This Google Apps Script helps pull transcripts for calls meeting pre-set filter criteria. In this particular script, the calls are filtered by date and call title keywords.

const API_KEY = "INSERT YOUR API ACCESS KEY HERE"; // ‼️ Insert your GONG API key here
const API_SECRET = "INSERT YOUR SECRET HERE"; // ‼️ Insert your GONG API secret here
const DOC_ID = "INSERT YOUR GOOGLE DOC ID HERE"; // ‼️ Insert your GOOGLE DOC ID here
const CALL_TITLE = "INSERT YOUR CALL TITLE KEYWORDS HERE"; // ‼️ Insert your keywords here, e.g., "Intro Meeting"


function onOpen() {
  var ui = DocumentApp.getUi();
  ui.createMenu("Gong API")
    .addItem("Refresh Transcripts", "fetchAndAppendAllTranscripts")
    .addToUi();
}

function fetchAndAppendAllTranscripts() {
  try {
    var doc = DocumentApp.openById(DOC_ID);
    var body = doc.getBody();

    var apiKey = API_KEY;
    var apiSecret = API_SECRET;
    var authHeader = "Basic " + Utilities.base64Encode(apiKey + ":" + apiSecret);

    var callIds = fetchAllCallIds(authHeader);
    if (callIds.length === 0) {
      Logger.log("⚠️ No matching calls found.");
      return;
    }

    var matchingCallIds = filterCallTitles(authHeader, callIds);
    if (matchingCallIds.length === 0) {
      Logger.log("⚠️ No calls match 'Sales Acceleration Weekly (Series)'.");
      return;
    }

    // Build the global speaker map once
    var globalSpeakerMap = buildGlobalSpeakerMap(authHeader, matchingCallIds);
    Logger.log("🎙️ Global Speaker Map created with " + Object.keys(globalSpeakerMap).length + " speakers");
    
    // DEBUG - Log the contents of the speaker map
    Logger.log("📋 SPEAKER MAP CONTENTS: " + JSON.stringify(globalSpeakerMap));

    // Use the global speaker map for all transcripts
    matchingCallIds.forEach(callId => fetchAndAppendTranscripts(callId, authHeader, body, globalSpeakerMap));
  } catch (error) {
    Logger.log("❌ Unexpected Error in fetchAndAppendAllTranscripts: " + error.message);
  }
}

function fetchAllCallIds(authHeader) {
  var scriptProperties = PropertiesService.getScriptProperties();
  var lastFetchedTime = scriptProperties.getProperty("lastFetchedTime") || "2025-01-01T00:00:00Z";  // ‼️ Adjust the start date
  var callSearchUrl = "https://us-13059.api.gong.io/v2/calls";
  var callIds = [];
  var cursor = null;
  var page = 1;
  var latestCallTimestamp = lastFetchedTime;

  do {
        var url = callSearchUrl + "?fromDateTime=" + encodeURIComponent(lastFetchedTime);
    if (cursor) {
      url += "&cursor=" + encodeURIComponent(cursor);
    }

    var response = UrlFetchApp.fetch(url, {
      method: "get",
      headers: { "Authorization": authHeader, "Content-Type": "application/json" },
      muteHttpExceptions: true
    });

    var jsonResponse = JSON.parse(response.getContentText());
    if (!jsonResponse.calls || jsonResponse.calls.length === 0) break;

    var newCallIds = jsonResponse.calls.map(call => call.id);
    callIds = callIds.concat(newCallIds);
    var fetchedTimestamps = jsonResponse.calls.map(call => call.startTime);
    var maxTimestamp = Math.max(...fetchedTimestamps);
    if (maxTimestamp > latestCallTimestamp) {
      latestCallTimestamp = new Date(maxTimestamp).toISOString();
    }

    cursor = jsonResponse.records?.cursor || null;
    page++;
  } while (cursor);

  if (callIds.length > 0) {
    scriptProperties.setProperty("lastFetchedTime", latestCallTimestamp);
  }

  return callIds;
}

function filterCallTitles(authHeader, callIds) {
  var matchingCallIds = [];
  var baseUrl = "https://us-13059.api.gong.io/v2/calls/extensive";

  for (var i = 0; i < callIds.length; i += 100) {
    var payload = JSON.stringify({
      "filter": { "callIds": callIds.slice(i, i + 100) },
      "contentSelector": { "exposedFields": { "metaData": true } }
    });

    var response = UrlFetchApp.fetch(baseUrl, {
      method: "post",
      headers: { "Authorization": authHeader, "Content-Type": "application/json" },
      payload: payload,
      muteHttpExceptions: true
    });

    var json = JSON.parse(response.getContentText());
    json.calls?.forEach(call => {
      if (call.metaData?.title?.toLowerCase().includes(CALL_TITLE)) {
        matchingCallIds.push(call.metaData.id);
      }
    });
  }

  return matchingCallIds;
}

function buildGlobalSpeakerMap(authHeader, callIds) {
  var globalSpeakerMap = {};
  
  // Process in batches of 100 calls
  for (var i = 0; i < callIds.length; i += 100) {
    var batchIds = callIds.slice(i, i + 100);
    var speakerUrl = "https://us-13059.api.gong.io/v2/calls/extensive";
    
    var speakerPayload = JSON.stringify({
      "filter": { "callIds": batchIds },
      "contentSelector": { "exposedFields": { "parties": true } }
    });

    var speakerOptions = {
      method: "post",
      headers: { "Authorization": authHeader, "Content-Type": "application/json" },
      payload: speakerPayload,
      muteHttpExceptions: true
    };

    var speakerResponse = UrlFetchApp.fetch(speakerUrl, speakerOptions);
    var speakerJson = JSON.parse(speakerResponse.getContentText());

    // Process each call's parties
    if (speakerJson.calls && speakerJson.calls.length > 0) {
      speakerJson.calls.forEach(call => {
        var parties = call.parties || [];
        parties.forEach(party => {
          if (party.speakerId != null && party.name) {
            // Ensure IDs are stored as strings
            globalSpeakerMap[String(party.speakerId)] = party.name;
          }
        });
      });
    }
    
    Logger.log("🎙️ Processed batch " + (i/100 + 1) + " for speaker mapping");
  }
  
  return globalSpeakerMap;
}

function fetchAndAppendTranscripts(callId, authHeader, body, speakerMap) {
  // Fetch Transcript
  var transcriptUrl = "https://us-13059.api.gong.io/v2/calls/transcript";
  var transcriptPayload = JSON.stringify({ "filter": { "callIds": [callId] } });

  var transcriptOptions = {
    method: "post",
    headers: { "Authorization": authHeader, "Content-Type": "application/json" },
    payload: transcriptPayload,
    muteHttpExceptions: true
  };

  var transcriptResponse = UrlFetchApp.fetch(transcriptUrl, transcriptOptions);
  var transcriptJson = JSON.parse(transcriptResponse.getContentText());

  // Log the first part of the transcript JSON to debug
  if (transcriptJson.callTranscripts && transcriptJson.callTranscripts.length > 0 && 
      transcriptJson.callTranscripts[0].transcript && transcriptJson.callTranscripts[0].transcript.length > 0) {
    var firstSpeaker = transcriptJson.callTranscripts[0].transcript[0];
    Logger.log("🔍 FIRST SPEAKER ID FORMAT: " + typeof firstSpeaker.speakerId + " - Value: " + firstSpeaker.speakerId);
  }

  // Fetch metadata for title, date, URL
  var metadataUrl = "https://us-13059.api.gong.io/v2/calls/extensive";
  var metadataPayload = JSON.stringify({
    "filter": { "callIds": [callId] },
    "contentSelector": { "exposedFields": { "metaData": true } }
  });

  var metadataOptions = {
    method: "post",
    headers: { "Authorization": authHeader, "Content-Type": "application/json" },
    payload: metadataPayload,
    muteHttpExceptions: true
  };

  var metadataResponse = UrlFetchApp.fetch(metadataUrl, metadataOptions);
  var metadataJson = JSON.parse(metadataResponse.getContentText());
  
  var metadata = metadataJson.calls?.[0]?.metaData;
  var rawDate = metadata?.scheduled ? new Date(metadata.scheduled) : new Date();
  var options = { year: 'numeric', month: 'long', day: 'numeric' };
  var formattedDate = rawDate.toLocaleDateString('en-US', options);
  var callURL = metadata?.url || "No URL available";
  
  if (!transcriptJson.callTranscripts || transcriptJson.callTranscripts.length === 0) {
    Logger.log("⚠️ No call transcripts found for call ID: " + callId);
    return;
  }

  var transcriptData = transcriptJson.callTranscripts[0]?.transcript || [];

  if (transcriptData.length === 0) {
    Logger.log("⚠️ No transcript data found for call ID: " + callId);
    return;
  }

  var transcriptText = "\n________________NEW CALL______________\n\n" +
    "\n📞 Call ID: " + callId +
    "\n📆 Call Date: " + formattedDate +
    "\n🔗 Call URL: " + callURL + "\n\n";

  // DEBUG - Check which IDs are in transcript vs. speaker map
  var transcriptIds = new Set();
  transcriptData.forEach(part => transcriptIds.add(String(part.speakerId)));
  Logger.log("🔍 TRANSCRIPT IDs: " + JSON.stringify(Array.from(transcriptIds)));
  Logger.log("🔍 SPEAKER MAP IDs: " + JSON.stringify(Object.keys(speakerMap)));

  transcriptData.forEach(function(part) {
    // Try different ID formats to ensure we find a match
    var speakerIdStr = String(part.speakerId);
    var speakerName = speakerMap[speakerIdStr];
    
    if (!speakerName) {
      Logger.log("⚠️ Missing speaker mapping for ID: " + speakerIdStr);
      speakerName = "Unknown Speaker";
    } else {
      // Debug the exact line being added
      var sampleLine = speakerName + ": " + (part.sentences[0]?.text || "");
      Logger.log("✏️ SAMPLE LINE: " + sampleLine);
    }
    
    part.sentences.forEach(function(sentence) {
      // Explicitly use the speaker name here to ensure it's properly applied
      var line = speakerName + ": " + sentence.text + "\n";
      transcriptText += line;
    });
  });

  if (transcriptText.trim() !== "") {
    body.appendParagraph(transcriptText);
    Logger.log("✅ Transcript added to Google Doc for call ID: " + callId);
  } else {
    Logger.log("⚠️ No transcript text created for call ID: " + callId);
  }
}
