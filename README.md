# ChatGPTApp Library

A library for interacting with the ChatGPT API. This library allows you to create chat conversations and call functions using the ChatGPT model.

## Installation

To use the ChatGPTApp library, follow these steps:

1. Open your Apps Script project in Google Workspace.
2. In the Apps Script editor, click on the menu **File > New > Script file**.
3. Name the new script file "ChatGPTApp".
4. Copy and paste the library code into the "ChatGPTApp.gs" file.
5. Save the script file.

## Usage

To use the ChatGPTApp library, follow the example below:

```javascript
// Initialize the ChatGPTApp library
const chatGPT = ChatGPTApp();

// Create a chat conversation
const chat = chatGPT.Chat();

// Add messages to the chat
chat.addMessage("Hello, how can I assist you?");

// Run the chat conversation
const response = chat.run();

// Process the response
console.log(response);
```

## ChatGPTApp API

### ChatGPTApp()

The `ChatGPTApp` object is the main entry point for using the ChatGPTApp library.

#### Properties

- `OpenAIKey`: A string representing the OpenAI API key.
- `GoogleKey`: A string representing the Google API key. This is only if you want to enable browsing.
- `BROWSING`: A boolean value indicating whether browsing is enabled.

### ChatGPTApp.Chat()

The `Chat` class represents a chat conversation.

#### Methods

- `addMessage(messageContent: string, system?: boolean): Chat`: Adds a message to the chat. The `messageContent` parameter is a string representing the content of the message. The `system` parameter is an optional boolean value indicating if the message is from the system. Returns the current `Chat` instance.
- `addFunction(functionObject: FunctionObject): Chat`: Adds a function to the chat. The `functionObject` parameter is an instance of the `FunctionObject` class. Returns the current `Chat` instance.
- `getMessages(): string[]`: Returns an array of strings representing the messages in the chat.
- `getFunctions(): FunctionObject[]`: Returns an array of `FunctionObject` instances representing the functions in the chat.
- `enableBrowsing(bool: boolean): Chat`: Enables or disables browsing in the chat. The `bool` parameter is a boolean value indicating whether browsing should be enabled. Returns the current `Chat` instance.
- `run(advancedParametersObject?: object): { functionName: string, functionArgs: JSON }`: Starts the chat conversation and returns the function called by the model, if any. The `advancedParametersObject` parameter is an optional object for advanced settings and specific usage only. Returns an object with the name of the function (`functionName`) and the arguments of the function (`functionArgs`).

### ChatGPTApp.FunctionObject()

The `FunctionObject` class represents a function known by the function calling model.

#### Methods

- `setName(newName: string): FunctionObject`: Sets the name for the function. Returns the current `FunctionObject` instance.
- `setDescription(newDescription: string): FunctionObject`: Sets the description for the function. Returns the current `FunctionObject` instance.
- `endWithResult(bool: boolean): FunctionObject`: Sets whether the conversation should automatically end when this function is called. Returns the current `FunctionObject` instance.
- `addParameter(name: string, newType: string, description: string, isOptional?: boolean): FunctionObject`: Adds a property (argument) to the function. The `name` parameter is the property name, `newType` is the property type, `description

is the property description, and `isOptional` is an optional boolean value indicating if the argument is required. Returns the current `FunctionObject` instance.
- `onlyReturnArguments(bool: boolean): FunctionObject`: Sets whether the conversation should automatically end when this function is called and return only the arguments in a stringified JSON object. Returns the current `FunctionObject` instance.
- `toJSON(): object`: Returns a JSON representation of the function.

## Examples

### Example 1: Adding Messages and Running the Chat Conversation

```javascript
const chatGPT = ChatGPTApp();

const chat = chatGPT.Chat();

// Add messages to the chat
chat.addMessage("Hello, how can I assist you?");
chat.addMessage("What's the weather like today?");

// Run the chat conversation
const response = chat.run();

console.log(response);
```

### Example 2: Adding Functions and Running the Chat Conversation

```javascript
const chatGPT = ChatGPTApp();

const chat = chatGPT.Chat();

// Create a function object
const calculateFunction = new chatGPT.FunctionObject()
  .setName("calculate")
  .setDescription("Perform a calculation")
  .addParameter("expression", "string", "The mathematical expression to calculate.");

// Add the function to the chat
chat.addFunction(calculateFunction);

// Add messages to the chat
chat.addMessage("Hello, how can I assist you?");
chat.addMessage("Please calculate the result of '2 + 2'.");

// Run the chat conversation
const response = chat.run();

console.log(response);

```

## Notes

- Ensure that you have a valid OpenAI API key and Google API key before using the library.
- Make sure to enable browsing if you need to perform web searches.
- The library is designed to work with the ChatGPT model and the function calling model. Ensure that you have the appropriate models enabled in your OpenAI account.
