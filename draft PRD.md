# Project Overview
GPT extension for Google Docs using Google Apps Script

# Core Functions
- Connect Google Apps Script to OpenAI API
- Add a Custom Menu "GPT Extension" in Google Docs
- Set Up the Sidebar UI that can provide a user-friendly interface where users enter prompts and view results
- There should be a setting panel, so one can edit the following parameters:
-- base url of openai compatible model
-- model itself (either chose from the short list of 4 most popular models or enter manually the name of the model)
-- temperature (from 0 to 1, of possible use progress bar or something like this)
-- max tokens (from 150 till infinite)

# Documentation
## Connect Google Apps Script to OpenAI API
'''
function setOpenAIKey() {
  const apiKey = 'your-api-key-here';
  PropertiesService.getScriptProperties().setProperty('OPENAI_API_KEY', apiKey);
}

// Function to make API calls to OpenAI
function callOpenAI(prompt) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  const apiUrl = 'https://api.openai.com/v1/chat/completions';
  
  const requestOptions = {
    'method': 'post',
    'headers': {
      'Authorization': 'Bearer ' + apiKey,
      'Content-Type': 'application/json'
    },
    'payload': JSON.stringify({
      'model': 'gpt-3.5-turbo',
      'messages': [{'role': 'user', 'content': prompt}],
      'temperature': 0.7,
      'max_tokens': 150
    })
  };

  try {
    const response = UrlFetchApp.fetch(apiUrl, requestOptions);
    const jsonResponse = JSON.parse(response.getContentText());
    return jsonResponse.choices[0].message.content;
  } catch (error) {
    Logger.log('Error: ' + error);
    return 'Error: ' + error;
  }
}

// Example function to test the API connection
function testOpenAI() {
  const prompt = "Hello, how are you?";
  const response = callOpenAI(prompt);
  Logger.log(response);
}
'''

## Add a Custom Menu "GPT Extension" in Google Docs
'''
// This function runs automatically when the document is opened
function onOpen() {
  const ui = DocumentApp.getUi();
  
  // Create the main menu
  ui.createMenu('GPT Extension')
    .addItem('Show Sidebar', 'showSidebar')
    .addSeparator()
    .addSubMenu(ui.createMenu('Quick Actions')
      .addItem('Summarize Selection', 'summarizeSelection')
      .addItem('Fix Grammar', 'fixGrammar')
      .addItem('Translate', 'translateText'))
    .addSeparator()
    .addItem('Settings', 'showSettings')
    .addToUi();
}

// Function to show sidebar
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('GPT Assistant');
  DocumentApp.getUi().showSidebar(html);
}

// Example function for summarizing selected text
function summarizeSelection() {
  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();
  
  if (selection) {
    const selectedText = selection.getRangeElements()[0].getElement().asText().getText();
    const prompt = `Please summarize the following text: ${selectedText}`;
    const summary = callOpenAI(prompt); // Using the previously defined OpenAI function
    
    // Display the result in a dialog
    const ui = DocumentApp.getUi();
    ui.alert('Summary', summary, ui.ButtonSet.OK);
  } else {
    DocumentApp.getUi().alert('Please select some text first.');
  }
}

// Example function for grammar correction
function fixGrammar() {
  // Similar implementation to summarizeSelection
}

// Example function for translation
function translateText() {
  // Similar implementation to summarizeSelection
}

// Function to show settings dialog
function showSettings() {
  const html = HtmlService.createHtmlOutputFromFile('Settings')
    .setWidth(400)
    .setHeight(300);
  DocumentApp.getUi().showModalDialog(html, 'Settings');
}
'''
## Settings panel
'''
<!DOCTYPE html>
<html>
  <head>
    <base target="_top">
    <style>
      .container {
        padding: 20px;
        font-family: Arial, sans-serif;
      }
      .form-group {
        margin-bottom: 15px;
      }
      label {
        display: block;
        margin-bottom: 5px;
        font-weight: bold;
      }
      input, select {
        width: 100%;
        padding: 8px;
        border: 1px solid #ddd;
        border-radius: 4px;
      }
      .range-container {
        display: flex;
        align-items: center;
        gap: 10px;
      }
      .range-value {
        min-width: 40px;
      }
      button {
        background-color: #4285f4;
        color: white;
        padding: 10px 20px;
        border: none;
        border-radius: 4px;
        cursor: pointer;
      }
      button:hover {
        background-color: #357abd;
      }
    </style>
  </head>
  <body>
    <div class="container">
      <div class="form-group">
        <label for="baseUrl">Base URL:</label>
        <input type="url" id="baseUrl" placeholder="https://api.openai.com/v1">
      </div>

      <div class="form-group">
        <label for="model">Model:</label>
        <select id="model">
          <option value="gpt-4">GPT-4</option>
          <option value="gpt-3.5-turbo">GPT-3.5 Turbo</option>
          <option value="gpt-3.5-turbo-16k">GPT-3.5 Turbo 16K</option>
          <option value="custom">Custom Model</option>
        </select>
      </div>

      <div class="form-group" id="customModelGroup" style="display: none;">
        <label for="customModel">Custom Model Name:</label>
        <input type="text" id="customModel" placeholder="Enter model name">
      </div>

      <div class="form-group">
        <label for="temperature">Temperature (0-1):</label>
        <div class="range-container">
          <input type="range" id="temperature" min="0" max="1" step="0.1" value="0.7">
          <span class="range-value" id="temperatureValue">0.7</span>
        </div>
      </div>

      <div class="form-group">
        <label for="maxTokens">Max Tokens:</label>
        <input type="number" id="maxTokens" min="150" value="150">
      </div>

      <button onclick="saveSettings()">Save Settings</button>
    </div>

    <script>
      // Show/hide custom model input based on selection
      document.getElementById('model').addEventListener('change', function() {
        const customModelGroup = document.getElementById('customModelGroup');
        customModelGroup.style.display = this.value === 'custom' ? 'block' : 'none';
      });

      // Update temperature value display
      document.getElementById('temperature').addEventListener('input', function() {
        document.getElementById('temperatureValue').textContent = this.value;
      });

      // Load existing settings when dialog opens
      google.script.run.withSuccessHandler(loadSettings).getSettings();

      function loadSettings(settings) {
        if (settings) {
          document.getElementById('baseUrl').value = settings.baseUrl || 'https://api.openai.com/v1';
          document.getElementById('model').value = settings.model || 'gpt-3.5-turbo';
          document.getElementById('temperature').value = settings.temperature || 0.7;
          document.getElementById('temperatureValue').textContent = settings.temperature || 0.7;
          document.getElementById('maxTokens').value = settings.maxTokens || 150;
          
          if (settings.model === 'custom' && settings.customModel) {
            document.getElementById('customModel').value = settings.customModel;
            document.getElementById('customModelGroup').style.display = 'block';
          }
        }
      }

      function saveSettings() {
        const settings = {
          baseUrl: document.getElementById('baseUrl').value,
          model: document.getElementById('model').value,
          customModel: document.getElementById('customModel').value,
          temperature: parseFloat(document.getElementById('temperature').value),
          maxTokens: parseInt(document.getElementById('maxTokens').value)
        };

        google.script.run
          .withSuccessHandler(closeDialog)
          .withFailureHandler(showError)
          .saveSettings(settings);
      }

      function closeDialog() {
        google.script.host.close();
      }

      function showError(error) {
        alert('Error saving settings: ' + error);
      }
    </script>
  </body>
</html>

And here's the corresponding Google Apps Script code to handle the settings:
// Get settings from Script Properties
function getSettings() {
  const properties = PropertiesService.getUserProperties();
  const settings = properties.getProperty('OPENAI_SETTINGS');
  return settings ? JSON.parse(settings) : null;
}

// Save settings to Script Properties
function saveSettings(settings) {
  const properties = PropertiesService.getUserProperties();
  properties.setProperty('OPENAI_SETTINGS', JSON.stringify(settings));
}

// Update the callOpenAI function to use the settings
function callOpenAI(prompt) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  const settings = getSettings() || {};
  
  const apiUrl = (settings.baseUrl || 'https://api.openai.com/v1') + '/chat/completions';
  const modelName = settings.model === 'custom' ? settings.customModel : settings.model;
  
  const requestOptions = {
    'method': 'post',
    'headers': {
      'Authorization': 'Bearer ' + apiKey,
      'Content-Type': 'application/json'
    },
    'payload': JSON.stringify({
      'model': modelName || 'gpt-3.5-turbo',
      'messages': [{'role': 'user', 'content': prompt}],
      'temperature': settings.temperature || 0.7,
      'max_tokens': settings.maxTokens || 150
    })
  };

  try {
    const response = UrlFetchApp.fetch(apiUrl, requestOptions);
    const jsonResponse = JSON.parse(response.getContentText());
    return jsonResponse.choices[0].message.content;
  } catch (error) {
    Logger.log('Error: ' + error);
    return 'Error: ' + error;
  }
}

'''

# Project File Structure
📁 GPT-Docs-Extension/
├── 📄 Code.gs            (Main script file with core functions)
├── 📄 Sidebar.html       (HTML for the main sidebar interface)
└── 📄 Settings.html      (HTML for settings dialog)
