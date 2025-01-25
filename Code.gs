SLACK_WEBHOOK = "XXXXXXX";
OPENAI_KEY = "XXXXXX";

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Meeting Recap')
    .addItem('Summarize & Send to Slack', 'summarizeAndSendToSlack')
    .addToUi();
}

function summarizeAndSendToSlack() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getRange("A:A").getValues().flat().filter(note => note.trim() !== "");

  if (data.length === 0) {
    SpreadsheetApp.getUi().alert("No notes found in Column A");
    return;
  }

  const summary = generateSummary(data.join("\n"));

  if (!summary) {
    SpreadsheetApp.getUi().alert("Failed to generate summary");
    return;
  }

  sendToSlack(summary);

  SpreadsheetApp.getUi().alert("Summary sent to Slack:\n\n" + summary);
}

function generateSummary(notes) {
  const url = "https://api.openai.com/v1/chat/completions";
  const payload = {
    model: "gpt-4",
    messages: [
      { role: "system", content: "You are a professional assistant that summarizes meeting and call notes into concise and actionable summarises" },
      { role: "user", content: `Here are my call notes:\n${notes}\n\nSummarize this into a brief and clear meeting recap.` }
    ],
    max_tokens: 150,
    temperature: 0.5
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: `Bearer ${OPENAI_KEY}`
    },
    payload: JSON.stringify(payload),
    mutteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());

    Logger.log(response);

    return json.choices[0]?.message?.content?.trim() || "No summary generated";
  } catch (error) {
    Logger.log("Error getting summary: " + error);
    return null;
  }
}

function sendToSlack(summary) {
  const payload = JSON.stringify({ text: summary });
  const options = {
    method: "post",
    contentType: "application/json",
    payload: payload
  };

  try {

    UrlFetchApp.fetch(SLACK_WEBHOOK, options);
    Logger.log("Summary sent to Slack");

  } catch (error) {
    Logger.log("Error sending to Slack: " + error);
  }
}
