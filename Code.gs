/**
 * The OpenAI API key constant
 * @constant {string}
 */
const OPENAI_API_KEY = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');

/**
 * Sets the OpenAI API key in script properties.
 * @param {string} apiKey - The OpenAI API key
 */
function setOpenAIKey(apiKey) {
  if (!apiKey) {
    throw new Error('API key cannot be empty');
  }
  PropertiesService.getScriptProperties().setProperty('OPENAI_API_KEY', OPENAI_API_KEY);
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
 * @param {Object} options - Optional parameters for the API call
 * @returns {string} The AI response text
 */
function callOpenAI(prompt, options = {}) {
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
      'model': options.model || settings.model || 'gpt-3.5-turbo',
      'messages': [{'role': 'user', 'content': prompt}],
      'temperature': options.temperature || settings.temperature || 0.7,
      'max_tokens': options.maxTokens || settings.maxTokens || 150
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

/**
 * Gets settings from Script Properties
 * @returns {Object|null} The settings object or null if not set
 */
function getSettings() {
  const properties = PropertiesService.getUserProperties();
  const settings = properties.getProperty('OPENAI_SETTINGS');
  return settings ? JSON.parse(settings) : null;
} 