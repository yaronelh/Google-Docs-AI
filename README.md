# Google Docs AI Writing Assistant

Want to use AI in Google Docs but prefer **OpenAI models or Claude's Sonnet** over Google's Gemini? Using this script You can achieve that by leveraging the **Google Apps Scripts** in Google Docs under the Tools menu. Simply add this script there to integrate AI-powered writing assistance directly into Google Docs using OpenAI's GPT API or Anthropic's Claude API.
This setup allows you to work seamlessly with either OpenAI or Claude models in your Google Docs. Just select the text, input a prompt in the side menu, and let the AI perform the changes with your preferred model.

## Features

- 🔄 Real-time text processing on selected text
- 💡 Multiple writing enhancement options
- 📝 Custom prompt interface
- 🎯 Precise text selection and replacement
- 🧰 Convenient sidebar interface

### Available Operations
- Custom Prompts to rewrite selected text
- Preset - Rephrase Selection
- Preset - Improve Writing
- Preset - Make Formal


```js
// Configuration
const OPENAI_API_KEY = 'YOUR_API_KEY_HERE'
const OPENAI_API_URL = 'https://api.openai.com/v1/chat/completions';

// Create main menu and context menu
function onOpen() {
  DocumentApp.getUi()
    .createMenu('AI helper v1')
    .addItem('Show Sidebar', 'showSidebar')
    .addItem('Rephrase Selection', 'rephraseSelection')
    .addItem('Improve Writing', 'improveWriting')
    .addItem('Make Formal', 'makeFormal')
    .addToUi();
}

// Get selected text accurately
function getSelectedText() {
  const selection = DocumentApp.getActiveDocument().getSelection();
  if (!selection) {
    DocumentApp.getUi().alert('Please select some text first.');
    return null;
  }

  let text = '';
  const elements = selection.getRangeElements();
  
  for (let i = 0; i < elements.length; i++) {
    const element = elements[i];
    const isPartial = element.isPartial();
    const textElement = element.getElement().asText();
    
    if (isPartial) {
      const startIndex = element.getStartOffset();
      const endIndex = element.getEndOffsetInclusive();
      text += textElement.getText().substring(startIndex, endIndex + 1);
    } else {
      text += textElement.getText();
    }
  }
  
  return {
    text: text,
    elements: elements
  };
}

// Replace selected text accurately
function replaceSelectedText(newText) {
  const selection = DocumentApp.getActiveDocument().getSelection();
  if (!selection) return;

  const elements = selection.getRangeElements();
  
  // Handle single element selection
  if (elements.length === 1) {
    const element = elements[0];
    const textElement = element.getElement().asText();
    
    if (element.isPartial()) {
      const startIndex = element.getStartOffset();
      const endIndex = element.getEndOffsetInclusive();
      const beforeText = textElement.getText().substring(0, startIndex);
      const afterText = textElement.getText().substring(endIndex + 1);
      textElement.setText(beforeText + newText + afterText);
    } else {
      textElement.setText(newText);
    }
  }
  // Handle multi-element selection
  else {
    // Replace only the selected portion in first element
    const firstElement = elements[0];
    if (firstElement.isPartial()) {
      const startIndex = firstElement.getStartOffset();
      const textElement = firstElement.getElement().asText();
      const beforeText = textElement.getText().substring(0, startIndex);
      textElement.setText(beforeText + newText);
    }
  }
}

// Main function to handle OpenAI API calls
function callOpenAI(text, instruction) {
  const prompt = instruction + text;
  
  const options = {
    'method': 'post',
    'headers': {
      'Authorization': 'Bearer ' + OPENAI_API_KEY,
      'Content-Type': 'application/json'
    },
    'payload': JSON.stringify({
      'model': 'GPT-4o',
      'messages': [{'role': 'user', 'content': prompt}],
      'temperature': 0.7,
      'max_tokens': 1000
    })
  };
  
  try {
    const response = UrlFetchApp.fetch(OPENAI_API_URL, options);
    const jsonResponse = JSON.parse(response.getContentText());
    return jsonResponse.choices[0].message.content.trim();
  } catch (error) {
    Logger.log('Error: ' + error);
    handleError(error);
    return null;
  }
}

// Function to rephrase selected text
function rephraseSelection() {
  const selection = getSelectedText();
  if (!selection) return;
  
  try {
    const improvedText = callOpenAI(selection.text, "Rephrase this text while maintaining its meaning: ");
    if (improvedText) {
      replaceSelectedText(improvedText);
    }
  } catch (error) {
    handleError(error);
  }
}

// Function to improve writing style
function improveWriting() {
  const selection = getSelectedText();
  if (!selection) return;
  
  try {
    const improvedText = callOpenAI(selection.text, "Improve this writing to make it more professional and clear: ");
    if (improvedText) {
      replaceSelectedText(improvedText);
    }
  } catch (error) {
    handleError(error);
  }
}

// Function to make text more formal
function makeFormal() {
  const selection = getSelectedText();
  if (!selection) return;
  
  try {
    const formalText = callOpenAI(selection.text, "Convert this text to a more formal tone: ");
    if (formalText) {
      replaceSelectedText(formalText);
    }
  } catch (error) {
    handleError(error);
  }
}

// Process custom prompt from sidebar
function processCustomPrompt(prompt) {
  const selection = getSelectedText();
  if (!selection) return;
  
  try {
    const improvedText = callOpenAI(selection.text, prompt + ": ");
    if (improvedText) {
      replaceSelectedText(improvedText);
    }
  } catch (error) {
    handleError(error);
  }
}

// Show sidebar
function showSidebar() {
  const html = HtmlService.createHtmlOutput(getSidebarHTML())
    .setTitle('AI Writing Assistant')
    .setWidth(300);
  DocumentApp.getUi().showSidebar(html);
}

// Sidebar HTML content
function getSidebarHTML() {
  return `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { font-family: Arial, sans-serif; margin: 10px; }
          .button { 
            background-color: #4285f4; 
            color: white; 
            padding: 8px 16px; 
            border: none; 
            border-radius: 4px; 
            cursor: pointer; 
            margin: 5px 0;
            width: 100%;
          }
          .button:hover { background-color: #357abd; }
          .input-area {
            width: 100%;
            height: 100px;
            margin: 10px 0;
            padding: 5px;
          }
          .section { margin: 15px 0; }
        </style>
      </head>
      <body>
        <div class="section">
          <h3>Selected Text Operations</h3>
          <button class="button" onclick="google.script.run.rephraseSelection()">Rephrase</button>
          <button class="button" onclick="google.script.run.improveWriting()">Improve</button>
          <button class="button" onclick="google.script.run.makeFormal()">Make Formal</button>
        </div>
        
        <div class="section">
          <h3>Custom Prompt</h3>
          <textarea id="customPrompt" class="input-area" placeholder="Enter your custom instruction..."></textarea>
          <button class="button" onclick="runCustomPrompt()">Execute</button>
        </div>

        <script>
          function runCustomPrompt() {
            const prompt = document.getElementById('customPrompt').value;
            google.script.run
              .withSuccessHandler(() => {
                document.getElementById('customPrompt').value = '';
              })
              .processCustomPrompt(prompt);
          }
        </script>
      </body>
    </html>
  `;
}

// Error handling function
function handleError(error) {
  Logger.log('Error: ' + error);
  DocumentApp.getUi().alert('An error occurred: ' + error.toString());
}


```



## Prerequisites

1. Google Account
2. OpenAI API Key ([Get it here](https://platform.openai.com/account/api-keys))
3. Google Docs (with permission to use Apps Script)

## Installation Guide

### Step 1: Open Google Docs
1. Open Google Docs
2. Create a new document or open an existing one
3. Go to `Tools > Script editor`

### Step 2: Set Up Apps Script
1. Delete any existing code in the script editor
2. Copy the entire code from `code.gs` in this repository
3. Paste it into the script editor
4. Replace `'YOUR_API_KEY_HERE'` with your actual OpenAI API key
5. Save the project (File > Save)
6. Give your project a name (e.g., "AI Writing Assistant")

### Step 3: Authorization
1. Close and reopen your Google Doc
2. Click on the new "AI Writing Helper" menu
3. When prompted, click "Continue" to grant necessary permissions
4. Select your Google Account
5. Click "Advanced" and then "Go to [Your Project Name] (unsafe)"
6. Click "Allow"

## Usage Instructions

### Using the Sidebar
1. Click "AI Writing Helper" > "Show Sidebar"
2. Select text in your document
3. Choose an operation from the sidebar:
   - Click "Rephrase" for paraphrasing
   - Click "Improve" for better writing
   - Click "Make Formal" for formal tone
   - Use custom prompt for specific instructions

### Using the Menu
1. Select text in your document
2. Click "AI Writing Helper" in the menu
3. Choose desired operation

## Best Practices

- Select text precisely - the tool works on exactly what you select
- Start with small sections to test functionality
- Set limit to your API key so it will not reach above what you're willing to pay

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.
