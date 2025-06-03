const url = 'https://api.openai.com/v1/chat/completions';
const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');

function GPT(prompt, range, model, temperature) {

  const payload = {
    model: model || 'gpt-4o-mini',
    messages: [
      { role: 'user', content: prompt += `\n\n${range}` }
      //, { role: 'system', content: system_prompt }
      ],
    temperature: temperature || 0.7
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + apiKey,
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    return json.choices ? json.choices[0].message.content.trim() : "Error: No response from OpenAI.";
  } catch (e) {
    return "Error: " + e.message;
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('AI Assistent')
    .addItem('Open Sidebar', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('AI Assistant');
  SpreadsheetApp.getUi().showSidebar(html);
}

function getActiveRangeReference() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getActiveRange();
  return `${sheet.getName()}!${range.getA1Notation()}`;
}
