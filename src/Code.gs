const ChatGPTApp = (function () {

  let OpenAIKey = "";

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
      let endingFunction = false;
      let onlyArgs = false;

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
       * OPTIONAL
       * If enabled, the conversation will automatically end when this function is called.
       * Default : false
       * @param {boolean} bool - Whether or not you wish for the option to be enabled. 
       * @returns {FunctionObject} - The current Function instance.
       */
      this.endWithResult = function (bool) {
        if (bool) {
          endingFunction = true;
        }
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
       * @param {boolean} [isOptional] - OPTIONAL - To set if the argument is required (default: true).
       * @returns {FunctionObject} - The current Function instance.
       */
      this.addParameter = function (name, newType, description, isOptional) {
        properties[name] = {
          type: newType,
          description: description
        };
        argumentsInRightOrder.push(name);
        if (!isOptional) {
          required.push(name);
        }
        return this;
      }

      /**
        * OPTIONAL
        * If enabled, the conversation will automatically end when this function is called and the chat will return the arguments in a stringified JSON object.
        * Default : false
        * @param {boolean} bool - Whether or not you wish for the option to be enabled. 
        * @returns {FunctionObject} - The current Function instance.
        */
      this.onlyReturnArguments = function (bool) {
        if (bool) {
          onlyArgs = true;
        }
        return this;
      }

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
          maximumNumberOfCalls: maximumNumberOfCalls,
          endingFunction: endingFunction,
          onlyArgs: onlyArgs
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
      let temperature = 0;
      let maxToken = 300;

      /**
       * Add a message to the chat.
       * @param {string} messageContent - The message to be added.
       * @param {boolean} [system] - OPTIONAL - True if message from system, False for user. 
       * @returns {Chat} - The current Chat instance.
       */
      this.addMessage = function (messageContent, system) {
        let role = "user";
        if (system) {
          role = "system";
        }
        messages.push({ role: role, content: messageContent });
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
       * Get the messages of the chat.
       * returns {string[]} - The messages of the chat.
       */
      this.getMessages = function () {
        return JSON.stringify(messages);
      };

      /**
       * Get the functions of the chat.
       * returns {FunctionObject[]} - The functions of the chat.
       */
      this.getFunctions = function () {
        return JSON.stringify(functions);
      };

      /**
       * Start the chat conversation.
       * Will return the chat answer.
       * If a function calling model is used, will call several functions until the chat decides that nothing is left to do.
       * @param {object} [advancedParametersObject] - OPTIONAL - For more advanced settings and specific usage only. {model, temperature, function_call}
       * @returns {Object} - the name (string) and arguments (JSON) of the function called by the model {functionName: name, functionArgs}
       */
      this.run = function (advancedParametersObject) {
        if (advancedParametersObject) {
          if (advancedParametersObject.hasOwnProperty("model")) {
            model = advancedParametersObject.model;
          }
          if (advancedParametersObject.hasOwnProperty("temperature")) {
            temperature = advancedParametersObject.temperature;
          }
        }

        let payload = {
          'messages': messages,
          'model': model,
          'max_tokens': maxToken,
          'temperature': temperature,
          'user': Session.getTemporaryActiveUserKey()
        };

        let functionCalling = false;
        if (advancedParametersObject) {
          if (advancedParametersObject.hasOwnProperty("function_calling")) { // the user has set a specific function to call
            Logger.log("coucou");
            payload.functions = functions;
            let function_calling = { name: advancedParametersObject.function_calling };
            payload.function_call = function_calling;
            functionCalling = true;
          }
        } else if (functions.length >> 0) { // the user has added functions, we enable function calling
          payload.functions = functions;
          payload.function_call = 'auto';
          functionCalling = true;
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
              'Authorization': 'Bearer ' + OpenAIKey
            },
            'payload': JSON.stringify(payload),
            'muteHttpExceptions': true
          };

          let response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', options);
          let responseCode = response.getResponseCode();

          if (responseCode === 200) {
            // The request was successful, exit the loop.
            responseMessage = JSON.parse(response.getContentText()).choices[0].message;
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
          return
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
            let functionArgs = parseResponse(responseMessage.function_call.arguments); // HERE THE JSON PARSE
            Logger.log(functionArgs);

            let argsOrder = [];
            let endWithResult = false;
            let onlyReturnArguments = false;

            for (let f in functions) {
              let currentFunction = functions[f].toJSON();
              if (currentFunction.name == functionName) {
                // get the args in the right order
                argsOrder = currentFunction.argumentsInRightOrder; // get the args in the right order
                endWithResult = currentFunction.endingFunction;
                onlyReturnArguments = currentFunction.onlyArgs;
                break;
              }
            }



            // Logger.log("argsOrder : " + argsOrder);
            // Logger.log("endWithResult : " + endWithResult);
            // Logger.log("onlyReturnArguments : " + onlyReturnArguments);

            if (endWithResult) {
              let functionResponse = callFunction(functionName, functionArgs, argsOrder);
              if (typeof functionResponse != "string") {
                functionResponse = String(functionResponse);
              }
              Logger.log({
                message: "Conversation stopped because end function has been called",
                functionName: functionName,
                functionArgs: functionArgs,
                functionResponse: functionResponse
              });
              return functionResponse;
            } else if (onlyReturnArguments) {
              Logger.log({
                message: "Conversation stopped because argument return has been enabled - No function has been called",
                functionName: functionName,
                functionArgs: functionArgs,
              });
              return functionArgs;
            } else {
              let functionResponse = callFunction(functionName, functionArgs, argsOrder);
              if (typeof functionResponse != "string") {
                functionResponse = String(functionResponse);
              }
              Logger.log({
                message: "Function calling called " + functionName,
                arguments: functionArgs,
              });
              // Inform the chat that the function has been called
              messages.push({
                "role": "assistant",
                "content": null,
                "function_call": { "name": functionName, "arguments": JSON.stringify(functionArgs) }
              });
              messages.push(
                {
                  "role": "function",
                  "name": functionName,
                  "content": functionResponse,
                }
              )
            }

            this.run();

          } else {
            // no function has been called 
            Logger.log({
              message: "No function has been called by the model",
            });
            // if no function has been found, stop here
            return responseMessage;
          }
        } else {
          // Logger.log(responseMessage.content)
          // Return the chat answer
          return responseMessage;
        }
      }
    }
  }

  function callFunction(functionName, jsonArgs, argsOrder) {
    // Parse JSON arguments
    var argsObj = jsonArgs;
    let argsArray = argsOrder.map(argName => argsObj[argName]);

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

  function parseResponse(response) {
    Logger.log("Input : " + response)
    try {
      let parsedReponse = JSON.parse(response);
      return parsedReponse;
    } catch (e) {
      // Split the response into lines
      let lines = response.trim().split('\n');

      if (lines[0] !== '{') {
        Logger.log('Unexpected start of response');
        return null;
      }
      else if (lines[lines.length - 1] !== '}') {
        lines.push('}');
      }
      for (let i = 1; i < lines.length - 1; i++) {
        let line = lines[i].trim();
        // For other lines, check for missing values or colons
        line = line.trimEnd();  // Strip trailing white spaces
        if (line[line.length - 1] !== ',') {
          if (line[line.length - 1] == ':') {
            // If line has the format "property":, add null,
            lines[i] = line + ' ""';
          } else if (!line.includes(':')) {
            lines[i] = line + ': ""';
          }
        }
      }
      // Reconstruct the response
      response = lines.join('\n').trim();

      // Try parsing the corrected response
      try {
        let parsedResponse = JSON.parse(response);
        return parsedResponse;
      } catch (e) {
        Logger.log(response)
        // If parsing still fails, log the error and return null.
        Logger.log('Error parsing corrected response: ' + e.message);
        return null;
      }
    }
  }

  return {
    /**
     * Create a new chat.
     * @params {string} apiKey - Your OPEN AI API key.
     * @returns {Chat} - A new Chat instance.
     */
    newChat: function () {
      return new Chat();
    },

    /**
     * Create a new function.
     * @returns {FunctionObject} - A new Function instance.
     */
    newFunction: function () {
      return new FunctionObject();
    },

    setAPIKey: function (apiKey) {
      OpenAIKey = apiKey;
    }
  }
}
)();