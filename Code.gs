// API and Authentication
const OPENAI_API_KEY = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');

/**
 * Sets the OpenAI API key in script properties.
 * @param {string} apiKey - The OpenAI API key
 */
function setOpenAIKey(apiKey) {
  if (!apiKey) {
    throw new Error('API key cannot be empty');
  }
  PropertiesService.getScriptProperties().setProperty('OPENAI_API_KEY', apiKey);
}

/**
 * Gets the stored OpenAI API key
 * @returns {string|null} The stored API key or null if not set
 */
function getOpenAIKey() {
  return PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
}

/**
 * Makes an API call to OpenAI's chat completions endpoint
 * @param {string} prompt - The user's input prompt
 * @returns {string} The AI response text
 */
function callOpenAI(prompt) {
  const apiKey = getOpenAIKey();
  if (!apiKey) {
    throw new Error('OpenAI API key not found. Please set it using setOpenAIKey()');
  }

  const settings = getSettings() || {};
  const baseUrl = (settings.baseUrl || 'https://api.openai.com/v1').replace(/\/$/, '');
  const apiUrl = `${baseUrl}/chat/completions`;
  
  const requestOptions = {
    'method': 'post',
    'headers': {
      'Authorization': 'Bearer ' + apiKey,
      'Content-Type': 'application/json'
    },
    'payload': JSON.stringify({
      'model': settings.model || 'gpt-3.5-turbo',
      'messages': [{'role': 'user', 'content': prompt}],
      'temperature': settings.temperature || 0.7,
      'max_tokens': settings.maxTokens || 150
    }),
    'muteHttpExceptions': true
  };

  try {
    const response = UrlFetchApp.fetch(apiUrl, requestOptions);
    const responseCode = response.getResponseCode();
    const contentText = response.getContentText();
    
    if (responseCode !== 200) {
      Logger.log(`API Error (${responseCode}): ${contentText}`);
      throw new Error(`API returned status code ${responseCode}`);
    }

    const jsonResponse = JSON.parse(contentText);
    return jsonResponse.choices[0].message.content;
  } catch (error) {
    Logger.log('Error calling OpenAI API: ' + error);
    throw new Error('Failed to get response from OpenAI: ' + error.message);
  }
}

/**
 * Creates the menu when the document is opened
 */
function onOpen() {
  const ui = DocumentApp.getUi();
  ui.createMenu('GPT Extension')
    .addItem('Show Sidebar', 'showSidebar')
    .addItem('Settings', 'showSettings')
    .addToUi();
}

/**
 * Shows the sidebar
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('GPT Assistant')
    .setWidth(300);
  DocumentApp.getUi().showSidebar(html);
}

/**
 * Shows the settings dialog
 */
function showSettings() {
  const html = HtmlService.createHtmlOutputFromFile('Settings')
    .setWidth(400)
    .setHeight(300);
  DocumentApp.getUi().showModalDialog(html, 'Settings');
}

/**
 * Gets settings from User Properties
 * @returns {Object} The settings object or default settings if not set
 */
function getSettings() {
  const properties = PropertiesService.getUserProperties();
  const settings = properties.getProperty('OPENAI_SETTINGS');
  return settings ? JSON.parse(settings) : {
    baseUrl: 'https://api.openai.com/v1',
    model: 'gpt-3.5-turbo',
    temperature: 0.7,
    maxTokens: 150
  };
}

/**
 * Saves settings to User Properties
 * @param {Object} settings - The settings object to save
 */
function saveSettings(settings) {
  if (!settings) {
    throw new Error('Settings object cannot be empty');
  }
  
  // Validate settings
  settings.temperature = parseFloat(settings.temperature) || 0.7;
  settings.maxTokens = parseInt(settings.maxTokens) || 150;
  settings.baseUrl = settings.baseUrl || 'https://api.openai.com/v1';
  settings.model = settings.model || 'gpt-3.5-turbo';

  const properties = PropertiesService.getUserProperties();
  properties.setProperty('OPENAI_SETTINGS', JSON.stringify(settings));
}

/**
 * Test function to verify API connection
 * @returns {string} The API response
 */
function testOpenAIConnection() {
  try {
    const prompt = "Hello, please respond with 'API connection successful' if you receive this message.";
    const response = callOpenAI(prompt);
    Logger.log('Test response: ' + response);
    return response;
  } catch (error) {
    Logger.log('Test failed: ' + error);
    throw error;
  }
}