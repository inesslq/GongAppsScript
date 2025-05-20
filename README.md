# Gong Apps Script
Google Apps Scripts to insert Gong call transcripts into Google Docs via Call ID or filters (e.g., name, date). Ideal for Sales and RevOps teams to store, review, and prep transcripts for AI analysis.

# Gong Transcript Inserters for Google Docs

These two Google Apps Scripts help you fetch and insert Gong call transcripts into a Google Doc. You can either provide a specific Call ID or use flexible filters like call title and date to find matching calls.

---

## Features

-  **Option 1: Prompt by Call ID**
    -  User is prompted to input a Gong Call ID
    -  Fetches and inserts the call transcript into the active Google Doc
-  **Option 2: Search by Filter**
      - Searches calls by speaker name, date, or other metadata (the sample script is using call title and date as filters)
      - Supports automatic transcript retrieval and speaker mapping
-  **Both scripts:**
      -  Match speaker names with transcript lines
      -  Pull call metadata (date, URL)
      -  Format the transcript clearly inside the Google Doc

---

## How to Use

1. Open a Google Doc
2. Go to **Extensions > Apps Script**
3. Paste the script and save it
4. Run the script by using the custom added menu in Google Docs UI

---

## APIs Used

- https://us-13059.api.gong.io/v2/calls/extensive — speaker info
- https://us-13059.api.gong.io/v2/calls/transcript — call transcript
- https://us-13059.api.gong.io/v2/calls — for filter-based fetching

---

## Output Example

**NEW CALL**

📞 Call ID: 123456789\
📆 Call Date: May 15, 2025\
🔗 Call URL: https://app.gong.io/call/...

Jane Smith: Thanks for joining.\
John Doe: Let’s start with Q1 results...
  
