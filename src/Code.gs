const ChatGPTApp = (function () {

  /**
   * @class
   * Class representing a chat message.
   */
  class Message {
    constructor() {
      let role = '';
      let content = '';

      /**
       * Set the role of the message.
       * @param {string} newRole - The role to be set.
       * @returns {Message} - The current Message instance.
       */
      this.setRole = function (newRole) {
        role = newRole;
        return this;
      };

      /**
       * Set the content of the message.
       * @param {string} newContent - The content to be set.
       * @returns {Message} - The current Message instance.
       */
      this.setContent = function (newContent) {
        content = newContent;
        return this;
      };

      /**
       * Returns a JSON representation of the message.
       * @returns {object} - The JSON representation of the message.
       */
      this.toJSON = function () {
        return {
          role: role,
          content: content
        };
      };
    }
  }

  /**
   * @class
   * Class representing a function known by function calling model
   */
  class FunctionObject {
    constructor() {
      let name = '';
      let description = '';
      let properties = {};
      let required = [];

      /**
       * Sets the name for a function.
       * @param {string} newName - The name to set for the function.
       * @returns {FunctionObject} - The current Function instance.
       */
      this.setName = function (newName) {
        name = newName;
        return this;
      };

      /**
       * Sets the description for a function.
       * @param {string} newDescription - The description to set for the function.
       * @returns {FunctionObject} - The current Function instance.
       */
      this.setDescription = function (newDescription) {
        description = newDescription;
        return this;
      };

      /**
       * Add a property (arg) of the function. 
       * Warning : required by default
       * check setPropertyAsUnrequired(propertyName)
       * @param {string} name - The property name.
       * @param {string} newType - The property type.
       * @param {string} description - The property description.
       * @returns {FunctionParameter} - The current FunctionParameter instance.
       */
      this.addParameter = function (name, newType, description) {
        properties[name] = {
          type: newType,
          description: description
        };
        required.push(name);
        return this;
      }

      /**
       * Sets a parameter as unrequired
       * @param {string} parameterName - The name of the unrequired parameter.
       * @returns {Function} - The current Function instance.
       */
      this.setPropertyAsUnrequired = function (parameterName) {
        var newRequiredArray = required.filter(function (element) {
          return element !== parameterName;
        })
        required = newRequiredArray;
        return this;
      };

      /**
       * Returns a JSON representation of the message.
       * @returns {object} - The JSON representation of the message.
       */
      this.toJSON = function () {
        return {
          name: name,
          description: description,
          parameters: {
            type: "object",
            properties: properties,
            required: required
          }
        };
      };
    }
  }

  /**
   * @class
   * Class representing a chat.
   */
  class Chat {
    constructor() {
      let messages = [];
      let functions = [];
      let model = "gpt-3.5-turbo"; // default 
      let apiKey = "";
      let temperature = 0;
      let maxToken = 300;

      /**
       * Add a message to the chat.
       * @param {JSON} message - The message to be added.Use JSON.stringify()
       * @returns {Chat} - The current Chat instance.
       */
      this.addMessage = function (message) {
        messages.push(message);
        return this;
      };

      /**
       * Add a function to the chat.
       * @param {FunctionObject} functionObject - The function to be added.
       * @returns {Chat} - The current Chat instance.
       */
      this.addFunction = function (functionObject) {
        functions.push(functionObject);
        return this;
      };

      /**
       * Set the open AI model (default : gpt-3.5-turbo)
       * @param {string} modelToUse - The model to use.
       * @returns {Chat} - The current Chat instance.
       */
      this.setModel = function (modelToUse) {
        model = modelToUse;
        return this;
      }

      /**
       * Set the open AI API key
       * @param {string} yourApiKey - Your API key.
       * @returns {Chat} - The current Chat instance.
       */
      this.setApiKey = function (yourApiKey) {
        apiKey = yourApiKey;
        return this;
      }

      /**
       * Set the temperature for the chat
       * @param {number} newTemperature - The temperature of the chat.
       * @returns {Chat} - The current Chat instance.
       */
      this.setTemperature = function (newTemperature) {
        temperature = newTemperature;
        return this;
      }

      /**
       * Set the maximum of token for one request.
       * @param {string} maximumNumberOfToken - The maximum number of token.
       * @returns {Chat} - The current Chat instance.
       */
      this.setMaxToken = function (maximumNumberOfToken) {
        maxToken = maximumNumberOfToken;
        return this;
      }

      /**
       * Get the messages of the chat.
       * @returns {Message[]} - The messages of the chat.
       */
      this.getMessages = function () {
        return JSON.stringify(messages);
      };

      /**
       * Get the functions of the chat.
       * @returns {FunctionObject[]} - The functions of the chat.
       */
      this.getFunctions = function () {
        return JSON.stringify(functions);
      };

      /**
       * Start the chat conversation.
       * Will return the chat answer.
       * If a function calling model is used, will call several functions until the chat decides that nothing is left to do.
       * @returns {Object} - the name (string) and arguments (JSON) of the function called by the model {functionName: name, functionArgs}
       */
      this.runConversation = function () {
        let functionCalling = false;
        let payload = {};
        if (model == "gpt-3.5-turbo-0613" || model == "gpt-4-0613") {
          payload = {
            'model': model,
            'messages': messages,
            'functions': functions,
            'function_call': 'auto',
          }
          functionCalling = true;
          Logger.log("Currently using function calling model");
        } else {
          payload = {
            'messages': messages,
            'model': model,
            'max_tokens': maxToken,
            'temperature': temperature,
            'user': Session.getTemporaryActiveUserKey()
          }
        }

        let options = {
          'method': 'post',
          'headers': {
            'Content-Type': 'application/json',
            'Authorization': 'Bearer ' + apiKey
          },
          'payload': JSON.stringify(payload),
        };
        let response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', options);
        let responseMessage = JSON.parse(response.getContentText()).choices[0].message;

        Logger.log({
          message: 'Got response from open AI API',
          response: JSON.stringify(responseMessage)
        });

        if (functionCalling) {
          // Check if GPT wanted to call a function
          if (responseMessage.function_call) {
            // Call the function
            let functionName = responseMessage.function_call.name;
            let functionArgs = responseMessage.function_call.arguments;

            callFunction(functionName, functionArgs);

            Logger.log({
              message: "Function calling called " + functionName,
              arguments: functionArgs
            });

            // Inform the chat that the function has been called
            const updateMessage = ChatGPTApp.newMessage().setRole("system").setContent("Function " + functionName + " has been executed with arguments " + functionArgs);
            this.addMessage(updateMessage);

            this.runConversation()

          } else {
            // no function has been called 
            Logger.log({
              message: "No function has been called by the model",
            });
            // if no function has been found, stop here
            return responseMessage;
          }
        } else {
          Logger.log(responseMessage.content)
          // Return the chat answer
          return responseMessage;
        }
      }
    }
  }

  function callFunction(functionName, jsonArgs) {
    // Parse JSON arguments
    var argsObj = JSON.parse(jsonArgs);
    var argsArray = Object.values(argsObj); // Get the values of the object properties

    // Call the function dynamically
    if (globalThis[functionName] instanceof Function) {
      globalThis[functionName].apply(null, argsArray);
    } else {
      Logger.log("Function not found or not a function: " + functionName);
    }
  }

  return {
    /**
     * Create a new chat.
     * @returns {Chat} - A new Chat instance.
     */
    newChat: function () {
      return new Chat();
    },

    /**
     * Create a new message.
     * @returns {Message} - A new Message instance.
     */
    newMessage: function () {
      return new Message();
    },

    /**
     * Create a new function.
     * @returns {FunctionObject} - A new Function instance.
     */
    newFunction: function () {
      return new FunctionObject();
    },
  }
}
)();