const ChatGPTApp = (function () {

  /**
   * @class
   * Class representing a chat message.
   */
  class Message {
    constructor(messageContent) {
      let role = "user";
      let content = messageContent;

      /**
       * Set the role of the message.
       * By default, the role is "user", you can set it as "system" by calling this function with param true
       * @param {boolean} bool - The role is "system"
       * @returns {Message} - The current Message instance.
       */
      this.setSystemInstruction = function (bool) {
        if (bool) {
            role = "system";
        }       
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
      let argumentsInRightOrder = [];
      let maximumNumberOfCalls = 1;

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
       * Sets the maximum of time the model will call a function.
       * Is here to avoid duplicated call.
       * Default value is 1.
       * @param {number} number - The number of time a function can be called.
       * @returns {FunctionObject} - The current Function instance.
       */
      this.setMaximumNumberOfCall = function (number) {
        maximumNumberOfCalls = number;
        return this;
      }

      /**
       * Add a property (arg) of the function. 
       * Warning : required by default
       * To set a parameter as unrequired, add false for the fourth argument.
       * Otherwise you can only use three arguments.
       * @param {string} name - The property name.
       * @param {string} newType - The property type.
       * @param {string} description - The property description.
       * @returns {Function} - The current Function instance.
       */
      this.addParameter = function (name, newType, description, isRequired) {
        isRequired = (isRequired === undefined) ? true : isRequired;
        properties[name] = {
          type: newType,
          description: description
        };
        argumentsInRightOrder.push(name);
        if (isRequired) {
          required.push(name);
        }
        return this;
      }

      // /**
      //  * Sets a parameter as unrequired
      //  * @param {string} parameterName - The name of the unrequired parameter.
      //  * @returns {Function} - The current Function instance.
      //  */
      // this.setPropertyAsUnrequired = function (parameterName) {
      //   var newRequiredArray = required.filter(function (element) {
      //     return element !== parameterName;
      //   })
      //   required = newRequiredArray;
      //   return this;
      // };

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
          },
          argumentsInRightOrder: argumentsInRightOrder,
          maximumNumberOfCalls: maximumNumberOfCalls
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
      let model = "gpt-3.5-turbo-0613"; // default 
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
       * @param {string} openAIKey - Your Open AI API key;
       * @returns {Object} - the name (string) and arguments (JSON) of the function called by the model {functionName: name, functionArgs}
       */
      this.runConversation = function (openAIKey) {
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

        let maxAttempts = 5;
        let attempt = 0;
        let success = false;

        let responseMessage;

        while (attempt < maxAttempts && !success) {
          let options = {
            'method': 'post',
            'headers': {
              'Content-Type': 'application/json',
              'Authorization': 'Bearer ' + openAIKey
            },
            'payload': JSON.stringify(payload)
            // 'muteHttpExceptions': true
          };

          let response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', options);
          let responseCode = response.getResponseCode();
          responseMessage = JSON.parse(response.getContentText()).choices[0].message;


          if (responseCode === 200) {
            // The request was successful, exit the loop.
            success = true;
          } else if (responseCode === 503) {
            // The server is temporarily unavailable, wait before retrying.
            let delay = Math.pow(2, attempt) * 1000; // Delay in milliseconds, starting at 1 second.
            Utilities.sleep(delay);
            attempt++;
          } else {
            // The request failed for another reason, log the error and exit the loop.
            console.error('Request failed with response code', responseCode);
            break;
          }
        }

        if (!success) {
          console.error('Failed to fetch the URL after', maxAttempts, 'attempts');
        }


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

            let argsOrder = [];
            // let numberOfTimeTheFunctionHasBeenCalled = 0;
            // let numberOfTimeTheFunctionCanBeCalled = 1;

            let functionCalled = [];

            for (let f in functions) {
              let currentFunction = functions[f].toJSON();
              if (currentFunction.name == functionName) {
                // get the args in the right order
                argsOrder = currentFunction.argumentsInRightOrder; // get the args in the right order

                // // count the functions calls
                // numberOfTimeTheFunctionCanBeCalled = currentFunction.maximumNumberOfCalls;
                // if (!functionCalled[functionName]) {
                //   functionCalled[functionName] = 0;
                // }
                // functionCalled[functionName]++;
                // numberOfTimeTheFunctionHasBeenCalled =  functionCalled[functionName];

                break;
              }
            }

            let functionResponse = callFunction(functionName, functionArgs, argsOrder);


            // if (numberOfTimeTheFunctionHasBeenCalled == numberOfTimeTheFunctionCanBeCalled) {
            //   Logger.log("reached maximum number of calls for the function.")
            //   var newFunctionsArray = functions.filter(function (element) {
            //     return element.name == functionName;
            //   })
            //   functions = newFunctionsArray;
            // }

            Logger.log({
              message: "Function calling called " + functionName,
              arguments: functionArgs,
              response: functionResponse
            });

            // Inform the chat that the function has been called
            messages.push({
                "role": "assistant",
                "content": null,
                "function_call": {"name": functionName, "arguments": functionArgs}}
            )
            messages.push(
              {
                "role": "function",
                "name": functionName,
                "content": functionResponse,
              }
            )

            // if (functions.length != 0) {
            this.runConversation(openAIKey);
            // } else {
            //   return;
            // }


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

  function callFunction(functionName, jsonArgs, argsOrder) {
    // Parse JSON arguments
    var argsObj = JSON.parse(jsonArgs);
    var argsArray = [];

    // Extract arguments in correct order
    for (let i = 0; i < argsOrder.length; i++) {
      let argName = argsOrder[i]; // get argument name from the argsOrder array
      argsArray.push(argsObj[argName]); // get value from argsObj
    }
    // Call the function dynamically
    if (globalThis[functionName] instanceof Function) {
      let functionResponse = globalThis[functionName].apply(null, argsArray);
      if (functionResponse) {
        // Logger.log(functionResponse);
        return functionResponse;
      } else {
        // Logger.log("no response");
        return "the function has been sucessfully executed but has nothing to return";
      }
    } else {
      Logger.log("Function not found or not a function: " + functionName);
      return "the function was not found and was not executed."
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
    newMessage: function (messageContent) {
      return new Message(messageContent);
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