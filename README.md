# ChatGPTApp Library Documentation

The `ChatGPTApp` library allows for the creation of chat applications with OpenAI's GPT models, by simplifying the creation of chat messages, function objects, and function parameters. This document outlines the structure of a typical conversation using this library.

## Initial Setup

Start by defining a new chat instance:

```javascript
let chat = ChatGPTApp.newChat();
```

This instance has various methods you can use to configure the chat, such as `setModel`, `setApiKey`, `setTemperature`, and `setMaxToken`.

```javascript
chat.setModel("gpt-3.5-turbo")
    .setApiKey("YOUR_OPENAI_API_KEY")
    .setTemperature(0.5)
    .setMaxToken(300);
```

## Adding Messages

Messages in a chat can be added with the `addMessage` method. To do this, first create a new `Message` instance:

```javascript
let message = ChatGPTApp.newMessage();
```

You can set the role and content of the message using the `setRole` and `setContent` methods respectively:

```javascript
message.setRole("system").setContent("Hello! How can I assist you today?");
```

Then, add this message to the chat:

```javascript
chat.addMessage(message);
```

## Adding Function Objects

Function objects represent the functions that the model can call during the conversation. 

### Creating FunctionParameter Instances

A `FunctionParameter` instance represents a parameter of the function. First, create a `FunctionParameter` instance:

```javascript
let functionParameter = ChatGPTApp.newFunctionParameter();
```

You can set the type of the function, and add properties (arguments) to it using `setType` and `addProperty` methods, respectively. Properties are added by specifying the name, type, and description. By default, all added properties are required. To make a property optional, use the `setPropertyAsUnrequired` method:

```javascript
functionParameter.setType("integer")
                .addProperty("number1", "integer", "First number to add.")
                .addProperty("number2", "integer", "Second number to add.")
                .setPropertyAsUnrequired("number2");
```

In this case, we have created a parameter object for a function that adds two numbers. The function requires at least one number (`number1`), and the second number (`number2`) is optional.

### Creating FunctionObject Instances

Next, create a new `FunctionObject`:

```javascript
let functionObject = ChatGPTApp.newFunction();
```

Set the name, description, and parameters of the function. The parameters are specified using the `FunctionParameter` instance created earlier:

```javascript
functionObject.setName("sum")
              .setDescription("Adds two numbers together.")
              .setParameters(functionParameter);
```

Here, we have defined a function named "sum" that adds two numbers together.

### Adding FunctionObject Instances to the Chat

Finally, add this function to the chat:

```javascript
chat.addFunction(functionObject);
```

Now, the chat application is aware of the "sum" function and it can be called during the conversation.

Keep in mind that the actual "sum" function must be implemented in your global scope, as the `ChatGPTApp` library will dynamically attempt to call this function when the conversation requires it.

## Running the Conversation

Once you've added all your messages and functions, you can start the conversation with the `runConversation` method:

```javascript
let response = chat.runConversation();
```

This method will run the conversation, call any functions if necessary, and return the final response from the chat.

Please note that the `runConversation` method takes into account whether a function calling model is used or not. In the case a function is called by the model, it will execute the function and continue the conversation until the chat decides that nothing is left to do. If no function is called, it simply returns the chat's response.

In case of function calling models, ensure that you have the functions defined in your global scope as the `callFunction` will try to call it dynamically based on the function name.

Please note that you need to replace `"YOUR_OPENAI_API_KEY"` with your actual OpenAI API key. Also, be sure to adjust the temperature and max tokens according to your needs. The model parameter should be one of the models supported by OpenAI.
