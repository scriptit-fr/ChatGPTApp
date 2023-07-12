# ChatGPTApp Google Apps Script Library Documentation

The ChatGPTApp is a library that facilitates the integration of OpenAI's GPT into your Google Apps Script projects. It allows for structured conversation, function calling and web browsing capabilities.

## Table of Contents

- [Setup](#setup)
- [Creating a New Chat](#creating-a-new-chat)
- [Adding Messages to the Chat](#adding-messages-to-the-chat)
- [Adding a Function to the Chat](#adding-a-function-to-the-chat)
- [Enabling Browsing](#enabling-browsing)
- [Running the Chat](#running-the-chat)
- [Function Object](#function-object)
- [Chat](#chat)
- [Note](#note)
- [Examples](#examples)

## Setup

To use the ChatGPTApp library, you first need to include the library code in your project. You then need to provide your OpenAI API key via `setOpenAIAPIKey()`.

If you wish to enable browsing capabilities, you will also need to provide your Google API key via `setGoogleAPIKey()`.

```javascript
ChatGPTApp.setOpenAIAPIKey("Your-OpenAI-API-Key");
ChatGPTApp.setGoogleSearchAPIKey("Your-Google-API-Key"); // if you want to enable browsing
```

To get an OpenAI API key, visit the OpenAI [website](https://openai.com/), create an account, and navigate to the API section to generate your API key.

To get a Google Custom Search API key (free) you can visit this [page](https://developers.google.com/custom-search/v1/introduction).

## Creating a New Chat

To start a new chat, call the `newChat()` method. This creates a new Chat instance.

```javascript
let chat = ChatGPTApp.newChat();
```

## Adding Messages to the Chat

You can add messages to your chat using the `addMessage()` method. Messages can be from the user or the system.

```javascript
chat.addMessage("Hello, how are you?");
chat.addMessage("Answer to the user in a professional way.", true);
```

## Adding a Function to the Chat

The `newFunction()` method allows you to create a new Function instance. You can then add this function to your chat using the `addFunction()` method.

```javascript
let functionObject = ChatGPTApp.newFunction()
  .setName("myFunction")
  .setDescription("This is a test function.")
  .addParameter("arg1", "string", "This is the first argument.");

chat.addFunction(functionObject);
```

From the moment that you add a function to chat, we will use Open AI's function calling features. 

For more information : [https://platform.openai.com/docs/guides/gpt/function-calling](https://platform.openai.com/docs/guides/gpt/function-calling) 

## Enabling Browsing

If you want to allow the model to perform web searches and fetch web pages, you can enable browsing.

```javascript
chat.enableBrowsing(true);
```

## Running the Chat

Once you have set up your chat, you can start the conversation by calling the `run()` method.

```javascript
let response = chat.run();
```

## Function Object

A `FunctionObject` represents a function that can be called by the chat.

Creating a function object and setting its name to the name of an actual function you have in your script will permit the library to call your real function.

### `setName(name)`

Sets the name of the function.

### `setDescription(description)`

Sets the description of the function.

### `addParameter(name, type, description, [isOptional])`

Adds a parameter to the function. Parameters are required by default. Set 'isOptional' to true to make a parameter optional.

### `endWithResult(bool)`

If enabled, the conversation with the chat will automatically end after this function is executed.

### `onlyReturnArguments(bool)`

If enabled, the conversation will automatically end when this function is called and the chat will return the arguments in a stringified JSON object.

### `toJSON()`

Returns a JSON representation of the function object.

## Chat

A `Chat` represents a conversation with the chat.

### `addMessage(messageContent, [system])`

Add a message to the chat. If `system` is true, the message is from the system, else it's from the user.

### `addFunction(functionObject)`

Add a function to the chat.

### `enableBrowsing(bool)`

Enable the chat to use a Google search engine to browse the web.

### `run([advancedParametersObject])`

Start the chat conversation. It sends all your messages and any added function to the chat GPT. It will return the last chat answer.

Supported attributes for the advanced parameters :

```javascript
advancedParametersObject = {
	temperature: temperature, 
	model: model,
	function_call: function_call
}
```

**Temperature** : Lower values for temperature result in more consistent outputs, while higher values generate more diverse and creative results. Select a temperature value based on the desired trade-off between coherence and creativity for your specific application.

**Model** : The OpenAI API is powered by a diverse [set of models](https://platform.openai.com/docs/models/overview) with different capabilities and price points. 

**Function_call** : If you want to force the model to call a specific function you can do so by setting `function_call: {"name": "<insert-function-name>"}`.

## Note

If you wish to disable the library logs and keep only your own, call `disableLogs()`:

```javascript
ChatGPTApp.disableLogs();
```

This can be useful for keeping your logs clean and specific to your application.

## Examples

### Example 1 : a simple chat gpt request

```javascript
 ChatGPTApp.setOpenAIAPIKey(OPEN_AI_API_KEY);

 const chat = ChatGPTApp.newChat();
 chat.addMessage("What are the steps to add an external library to my Google Apps Script project?");

 const chatAnswer = chat.run();
 Logger.log(chatAnswer);
```

### Example 2 : automatically create a draft reply for the last email in Gmail inbox

```javascript
 ChatGPTApp.setOpenAIAPIKey(OPEN_AI_API_KEY);
 const chat = ChatGPTApp.newChat();

 // Fonction calling : rédaction automatique de brouillons pour les mails gmails 

 var getLatestThreadFunction = ChatGPTApp.newFunction()
    .setName("getLatestThread")
    .setDescription("Retrieve information from the last message received.");

 var createDraftResponseFunction = ChatGPTApp.newFunction()
    .setName("createDraftResponse")
    .setDescription("Create a draft response.")
    .addParameter("threadId", "string", "the ID of the thread to retrieve")
    .addParameter("body", "string", "the body of the email in plain text");

  var resp = ChatGPTApp.newChat()
    .addMessage("You are an assistant managing my Gmail inbox.", true)
    .addMessage("Retrieve the latest message I received and draft a response.")
    .addFunction(getLatestThreadFunction)
    .addFunction(createDraftResponseFunction)
    .run();

  console.log(resp);
```

### Example 3 : use case for onlyReturnArgument()

```javascript
const ticket = "Hello, could you check the status of my subscription under customer@example.com";

  chat.addMessage("You just received this ticket : " + ticket);
  chat.addMessage("What's the customer email address ? You will give it to me using the function getEmailAddress.");

  const myFunction = ChatGPTApp.newFunction() // in this example, getEmailAddress is not actually a real function in your script
    .setName("getEmailAddress")
    .setDescription("To give the user an email address")
    .addParameter("emailAddress", "string", "the email address")
    .onlyReturnArguments(true) // you will get your parameters in a json object

  chat.addFunction(myFunction);

  const chatAnswer = chat.run();
  Logger.log(chatAnswer["emailAddress"]); // the name of the parameter of your "fake" function

  // output : 	"customer@example.com"
```

### Example 4 : browsing

```javascript
 const message = "You're a google support agent, a customer is asking you how to install a library he found on github in a google appscript project."

 const chat = ChatGPTApp.newChat();
 chat.addMessage(message);
 chat.addMessage("Browse this website to answer : https://developers.google.com/apps-script", true)
 chat.enableBrowsing(true);

 const chatAnswer = chat.run();
 Logger.log(chatAnswer);
```
