# ChatGPTApp Documentation

ChatGPTApp is a library that allows for easier usage of the OpenAI API with Google Apps Script. It provides a simple interface for creating and managing chat sessions with the OpenAI API and integrating it into your Google Apps Script projects.

## ChatGPTApp API

The following classes and methods are available with the ChatGPTApp library:

### ChatGPTApp.newChat()

This function creates and returns a new instance of the `Chat` class. This class represents a conversation.

```javascript
let myChat = ChatGPTApp.newChat();
```

### ChatGPTApp.newMessage(messageContent)

This function creates and returns a new instance of the `Message` class. It requires a string representing the message content as a parameter.
By default, the message is assigned the role "user" (most common use case). You will see bellow how you can modify it. 

```javascript
let myMessage = ChatGPTApp.newMessage("Hello, how are you?");
```

### ChatGPTApp.newFunction()

This function creates and returns a new instance of the `FunctionObject` class. 

```javascript
let myFunction = ChatGPTApp.newFunction();
```

## Classes

### Message

#### Message.setSystemInstruction(bool)

Sets the role of the message as 'system' if `bool` is `true`. If `false`, the role remains 'user'.
The use of this function is of course optionnal.

```javascript
myMessage.setSystemInstruction(true);
```

#### Message.setContent(newContent)

Modifies the content of the message.

```javascript
myMessage.setContent("What's the weather today?");
```

### FunctionObject

#### FunctionObject.setName(newName)

Sets the name of the function.

```javascript
myFunction.setName("getWeather");
```

#### FunctionObject.setDescription(newDescription)

Sets the description of the function.

```javascript
myFunction.setDescription("Gets the current weather.");
```

#### FunctionObject.addParameter(name, newType, description, isRequired)

Adds a property or argument to the function. The fourth parameter, `isRequired`, is optional and defaults to `true` if not specified.

```javascript
myFunction.addParameter("location", "string", "The location to get the weather for", true);
```

#### FunctionObject.toJSON()

Returns a JSON representation of the function.

```javascript
let functionJSON = myFunction.toJSON();
```

### Chat

#### Chat.addMessage(message)

Adds a message to the chat. The parameter `message` should be an instance of the `Message` class.

```javascript
myChat.addMessage(myMessage);
```

#### Chat.addFunction(functionObject)

Adds a function to the chat. The parameter `functionObject` should be an instance of the `FunctionObject` class.

```javascript
myChat.addFunction(myFunction);
```

#### Chat.getMessages()

Returns the messages of the chat in a JSON string format.

```javascript
let messages = myChat.getMessages();
```

#### Chat.getFunctions()

Returns the functions of the chat in a JSON string format.

```javascript
let functions = myChat.getFunctions();
```

#### Chat.runConversation(openAIKey, advancedParametersObject)

Starts the chat conversation and returns the model's response. The `advancedParametersObject` parameter is optional and can be used to specify advanced parameters such as `model`, `temperature`, and `function_calling`.

```javascript
let response = myChat.runConversation('your_openai_key', { model: "gpt-3.5-turbo", temperature: 0.5 });
```
