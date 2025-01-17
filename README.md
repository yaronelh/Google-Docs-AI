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
const OPENAI_API_KEY = 'YOUR_API_KEY_HERE';
const OPENAI_API_URL = 'https://api.openai.com/v1/chat/completions';

// Create main menu and context menu
function onOpen() {
  DocumentApp.getUi()
    .createMenu('AI Writing Assistant')
    .addItem('Show Sidebar', 'showSidebar')
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
function callOpenAI(context, userInstruction, selectedText) {
  const data = {
    model: 'gpt-4o',
    messages: [
      { 
        role: 'system', 
        content: context 
      },
      {
        role: 'user',
        content: `${userInstruction}`
      }
    ],

    temperature: 0.7,
    max_tokens: 1000,

    prediction: {
      type: 'content',
      content: selectedText
    }
  };

  const options = {
    method: 'post',
    headers: {
      Authorization: 'Bearer ' + OPENAI_API_KEY,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(data)
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


// Function to improve writing with context
function improveWritingWithContext(context) {
  const selection = getSelectedText();
  if (!selection) return;

  try {
    const prompt = `Go over the text, check it for grammatical and punctuation errors and correct them, maintain as much of the text as possible to preserve the tone of voice, and make sure to not change the meaning of the text:\n\n${selection.text}`;
    const improvedText = callOpenAI(context, prompt, selection.text);
    if (improvedText) {
      replaceSelectedText(improvedText);
    }
  } catch (error) {
    handleError(error);
  }
}

// Process custom prompt with context
function processCustomPromptWithContext(context, customPrompt) {
  const selection = getSelectedText();
  if (!selection) return;

  try {
    const prompt = `${customPrompt}:\n\n${selection.text}`;
    const improvedText = callOpenAI(context, prompt, selection.text);
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
      body {
        font-family: "Roboto", Arial, sans-serif; 
        margin: 12px;
        background-color: #ffffff;
        color: #202124;
      }

      h3 {
        color: #185abc;
        margin-bottom: 6px;
        font-weight: 500;
      }

      .description {
        color: #5f6368;
        font-size: 13px;
        margin-bottom: 10px;
      }

      .button {
        background-color: #4285f4; 
        color: white; 
        padding: 8px 16px; 
        border: none; 
        border-radius: 4px; 
        cursor: pointer; 
        margin: 5px 0;
        width: 100%;
        font-size: 14px;
      }

      .button:hover {
        background-color: #357abd;
      }

      .input-area {
        width: 100%;
        box-sizing: border-box;
        padding: 5px;
        height: 180px;
        font-size: 13px;
        border: 1px solid #dadce0;
        border-radius: 4px;
        margin-bottom: 8px;
      }

      .section {
        margin-top: 16px;
      }

      .section h3 {
        margin-top: 0; 
      }

      .divider {
        border: none; /* Removes default border */
        height: 1px; /* Thickness of the line */
        background-color: #dadce0; /* Light gray color for the line */
        width: 100%; /* Full width */
        margin: 12px 0; /* Top and bottom margin of 12px */
      }
    </style>
  </head>
  <body>
    <div class="section">
      <div class="description">This tool helps you adjust your selected text using OpenAI's API, and is made for those who prefer to use other models besides Google Gemini.</div>
      <h3>Context</h3>
      <div class="description">The context is usually referred to as the 'system' section of the prompt.</div>
      <textarea id="contextInput" class="input-area" placeholder="Enter context/system prompt here..."></textarea>
    </div>

    <div class="section">
      <h3>Custom Prompt</h3>
      <div class="description">Write your own custom prompt to run on the selected text</div>
      <textarea id="customPrompt" class="input-area" placeholder="Enter your custom instruction..."></textarea>
      <div class="description">Clicking 'Run' will execute your context and custom prompt on the selected text.</div>
      <button class="button" onclick="runCustomPrompt()">Run</button>
    </div>
    <hr class="divider">
    <div class="section">
      <h3>Presets</h3>
      <div class="description">Uses your 'context' with a preset to refine the selected text.</div>
      <button class="button" onclick="runImprove()">Improve Selected Text</button>
    </div>

    <script>
      function getContextAndPrompt() {
        const context = document.getElementById('contextInput').value || '';
        const customPrompt = document.getElementById('customPrompt').value || '';
        return { context, customPrompt };
      }

      function runImprove() {
        const { context } = getContextAndPrompt();
        google.script.run.improveWritingWithContext(context);
      }

      function runCustomPrompt() {
        const { context, customPrompt } = getContextAndPrompt();
        if (customPrompt) {
          google.script.run.processCustomPromptWithContext(context, customPrompt);
        } else {
          alert('Please enter a custom prompt.');
        }
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
