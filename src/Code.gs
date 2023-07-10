const ChatGPTApp = (function () {

  let OpenAIKey = "";
  let GoogleCustomSearchAPIKey = "";
  let BROWSING = false;

  let SEARCH_RESULTS = [];

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
       * Adds a property (arg) to the function.
       * 
       * Note: Parameters are required by default. Set 'isOptional' to true to make a parameter optional.
       *
       * @param {string} name - The property name.
       * @param {string} type - The property type.
       * @param {string} description - The property description.
       * @param {boolean} [isOptional] - To set if the argument is optional (default: false).
       * @returns {FunctionObject} - The current Function instance.
       */
      this.addParameter = function (name, type, description, isOptional = false) {
        let itemsType;


        if (String(type).includes("Array")) {
          let startIndex = type.indexOf("<") + 1;
          let endIndex = type.indexOf(">");
          itemsType = type.slice(startIndex, endIndex);
          type = "array";
        }

        properties[name] = {
          type: type,
          description: description
        };

        if (type === "array") {
          if (itemsType) {
            properties[name]["items"] = {
              type: itemsType
            }
          } else {
            Logger.log("Please precise the type of the items contained in the array when calling addParameter. Use format Array.<itemsType> for the type parameter.");
            return
          }

        }

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

  let webSearchFunction = new FunctionObject()
    .setName("webSearch")
    .setDescription("Perform a web search via the Google Custom Search JSON API. Returns an array of search results (including the URL, title and text snippets for each result)")
    .addParameter("q", "string", "the query for the web search.");

  let urlFetchFunction = new FunctionObject()
    .setName("urlFetch")
    .setDescription("Fetch the viewable contents of a web page. It will strip HTML tags, returning just raw text.")
    .addParameter("url", "string", "The URL to fetch.");

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

      this.enableBrowsing = function (bool) {
        if (bool) {
          BROWSING = true
        }
        return this;
      }

      /**
       * Start the chat conversation.
       * Will return the chat answer.
       * If a function calling model is used, will call several functions until the chat decides that nothing is left to do.
       * @param {object} [advancedParametersObject] - OPTIONAL - For more advanced settings and specific usage only. {model, temperature, function_call}
       * @returns {Promise<string>} - the stringifed last message of the chat 
       */
      this.run = async function (advancedParametersObject) {
        if (advancedParametersObject) {
          if (advancedParametersObject.hasOwnProperty("model")) {
            model = advancedParametersObject.model;
          }
          if (advancedParametersObject.hasOwnProperty("temperature")) {
            temperature = advancedParametersObject.temperature;
          }
        }

        if (BROWSING && messages[messages.length - 1].role !== "function") {
          messages.push({ role: "system", content: "You are able to perform queries on Google search using the function webSearch, then open results and get the content of a web page using the function urlFetch." });
          functions.push(webSearchFunction);
          functions.push(urlFetchFunction);
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
        let endReason;

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
            endReason = JSON.parse(response.getContentText()).choices[0].finish_reason;
            if (endReason == "length") {
              Logger.log({
                message: "WARNING : Answer has been troncated because it was too long. To resolve this issue, you can increase the max_tokens property",
                currentMaxToken: maxToken
              });
            }
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
          return messages[-1];
        }

        // Logger.log({
        //   message: 'Got response from open AI API',
        //   response: JSON.stringify(responseMessage)
        // });

        if (functionCalling) {
          // Check if GPT wanted to call a function
          if (responseMessage.function_call) {
            // Call the function
            let functionName = responseMessage.function_call.name;
            let functionArgs = parseResponse(responseMessage.function_call.arguments);

            let argsOrder = [];
            let endWithResult = false;
            let onlyReturnArguments = false;

            for (let f in functions) {
              let currentFunction = functions[f].toJSON();
              if (currentFunction.name == functionName) {
                argsOrder = currentFunction.argumentsInRightOrder; // get the args in the right order
                endWithResult = currentFunction.endingFunction;
                onlyReturnArguments = currentFunction.onlyArgs;
                break;
              }
            }

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
              return messages[messages.length - 1];


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
              if (functionName !== "webSearch" && functionName !== "urlFetch") {
                Logger.log({
                  message: "Function calling called " + functionName,
                  arguments: functionArgs,
                });
              }
              else if (functionName == "webSearch") {
                SEARCH_RESULTS = JSON.parse(functionResponse);
              }
              else if (functionName == "urlFetch") {
                if (!functionResponse) {
                  Logger.log("The website didn't respond, going back to the results page");
                  let searchResults = JSON.parse(messages[messages.length - 1].content);
                  let newSearchResult = searchResults.filter(function (obj) {
                    return obj.link !== functionArgs.url;
                  });
                  messages[messages.length - 1].content = JSON.stringify(newSearchResult);
                }

              }
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

          }
          else {
            // no function has been called 
            Logger.log({
              message: "No function has been called by the model",
            });
            // if no function has been found, stop here
            return messages[messages.length - 1];
          }
        }
        else {
          return messages[messages.length - 1];
        }
      }
    }
  }

  function callFunction(functionName, jsonArgs, argsOrder) {
    // Handle internal functions
    if (functionName == "webSearch") {
      return webSearch(jsonArgs.q);
    }
    if (functionName == "urlFetch") {
      return urlFetch(jsonArgs.url);
    }
    // Parse JSON arguments
    var argsObj = jsonArgs;
    let argsArray = argsOrder.map(argName => argsObj[argName]);

    // Call the function dynamically
    if (globalThis[functionName] instanceof Function) {
      let functionResponse = globalThis[functionName].apply(null, argsArray);
      if (functionResponse) {
        return functionResponse;
      } else {
        return "the function has been sucessfully executed but has nothing to return";
      }
    } else {
      Logger.log("Function not found or not a function: " + functionName);
      return "the function was not found and was not executed."
    }
  }

  function parseResponse(response) {
    try {
      let parsedReponse = JSON.parse(response);
      return parsedReponse;
    } catch (e) {
      // Split the response into lines
      let lines = response.trim().split('\n');

      if (lines[0] !== '{') {
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
          } else if (line[line.length - 1] !== '"') {
            lines[i] = line + '"';
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
        // If parsing still fails, log the error and return null.
        console.warn({
          message: 'Error parsing corrected response: ' + e.message,
          argumentJSON: response
        });
        return null;
      }
    }
  }

  function webSearch(q) {
    Logger.log(`Web search : "${q}"`);
    const searchEngineId = "221c662683d054b63";
    const url = `https://www.googleapis.com/customsearch/v1?key=${GoogleCustomSearchAPIKey}&cx=${searchEngineId}&q=${encodeURIComponent(q)}`;

    let response = UrlFetchApp.fetch(url);
    let data = JSON.parse(response.getContentText());

    let resultsInfo = [];
    let resultsContent = [];

    if (data.items) {
      resultsInfo = data.items.map(function (item) {
        return {
          title: item.title,
          link: item.link,
          snippet: item.snippet
        };
      }).filter(Boolean); // Remove undefined values from the results array
    }
    return JSON.stringify(resultsInfo);
  }

  function urlFetch(url) {
    Logger.log("Clicked on a link");
    let pageContent = UrlFetchApp.fetch(url).getContentText();
    pageContent = sanitizeHtml(pageContent);
    return pageContent;
  }

  function sanitizeHtml(pageContent) {

    var startBody = pageContent.indexOf('<body');
    var endBody = pageContent.indexOf('</body>') + '</body>'.length;
    var bodyContent = pageContent.slice(startBody, endBody);

    // naive way to extract paragraphs
    var paragraphs = bodyContent.match(/<p[^>]*>([^<]+)<\/p>/g);

    var usefullContent = "";
    if (paragraphs) {
      for (var i = 0; i < paragraphs.length; i++) {
        var paragraph = paragraphs[i];

        // remove HTML tags
        var taglessParagraph = paragraph.replace(/<[^>]+>/g, '');

        usefullContent += taglessParagraph + '\n';
      }
    }

    // remove doctype
    usefullContent = usefullContent.replace(/<!DOCTYPE[^>]*>/i, '');

    // remove html comments
    usefullContent = usefullContent.replace(/<!--[\s\S]*?-->/g, '');

    // remove new lines
    usefullContent = usefullContent.replace(/\n/g, '');

    // replace multiple spaces with a single space
    usefullContent = usefullContent.replace(/ +(?= )/g, '');

    // trim the result
    usefullContent = usefullContent.trim();

    return usefullContent;
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

    setOpenAIAPIKey: function (apiKey) {
      OpenAIKey = apiKey;
    },

    setGoogleAPIKey: function (apiKey) {
      GoogleCustomSearchAPIKey = apiKey;
    },
  }
}
)();