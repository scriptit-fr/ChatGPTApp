/*
 ChatGPTApp
 https://github.com/scriptit-fr/ChatGPTApp
 
 Copyright (c) 2023 Guillemine Allavena - Romain Vialard
 
 Licensed under the Apache License, Version 2.0 (the "License");
 you may not use this file except in compliance with the License.
 You may obtain a copy of the License at
 
 http://www.apache.org/licenses/LICENSE-2.0
 
 Unless required by applicable law or agreed to in writing, software
 distributed under the License is distributed on an "AS IS" BASIS,
 WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 See the License for the specific language governing permissions and
 limitations under the License.
 */


const ChatGPTApp = (function () {

  let OpenAIKey = "";
  let GoogleCustomSearchAPIKey = "";
  let ENABLE_LOGS = true;
  let SITE_SEARCH;
  let KNOWLEDGE_LINK;

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
      let endingFunction = false;
      let onlyArgs = false;

      /**
       * Sets the name of a function.
       * @param {string} nameOfYourFunction - The name to set for the function.
       * @returns {FunctionObject} - The current Function instance.
       */
      this.setName = function (nameOfYourFunction) {
        name = nameOfYourFunction;
        return this;
      };

      /**
       * Sets the description of a function.
       * @param {string} descriptionOfYourFunction - The description to set for the function.
       * @returns {FunctionObject} - The current Function instance.
       */
      this.setDescription = function (descriptionOfYourFunction) {
        description = descriptionOfYourFunction;
        return this;
      };

      /**
       * OPTIONAL
       * If enabled, the conversation with the chat will automatically end when this function is called.
       * Default : false, eg the function is sent to the chat that will decide what the next action shoud be accordingly. 
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
       * Adds a property (an argument) to the function.
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
            console.log("Please precise the type of the items contained in the array when calling addParameter. Use format Array.<itemsType> for the type parameter.");
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
      let temperature = 0.5;
      let maxToken = 300;
      let browsing = false;

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
       * OPTIONAL
       * 
       * Enable the chat to use a google serach engine to browse the web.
       * @param {boolean} bool - Whether or not you wish for the option to be enabled. 
       * @returns {Chat} - The current Chat instance.
       */
      this.enableBrowsing = function (bool) {
        if (bool) {
          browsing = true
        }
        return this;
      }

      /**
       * Start the chat conversation.
       * Sends all your messages and eventual function to chat GPT.
       * Will return the last chat answer.
       * If a function calling model is used, will call several functions until the chat decides that nothing is left to do.
       * @param {{model?: "gpt-3.5-turbo" | "gpt-3.5-turbo-16k" | "gpt-4" | "gpt-4-32k", temperature?: number, maxToken?: number, function_call?: string}} [advancedParametersObject] - OPTIONAL - For more advanced settings and specific usage only. {model, temperature, function_call}
       * @returns {object} - the last message of the chat 
       */
      this.run = function (advancedParametersObject) {
        if (advancedParametersObject) {
          if (advancedParametersObject.model) {
            model = advancedParametersObject.model;
          }
          if (advancedParametersObject.temperature) {
            temperature = advancedParametersObject.temperature;
          }
          if (advancedParametersObject.maxToken) {
            maxToken = advancedParametersObject.maxToken;
          }
        }

        if (KNOWLEDGE_LINK) {
          let knowledge = urlFetch(KNOWLEDGE_LINK);
          messages.push({ role: "system", content: `Here's an article to help:\n\n${knowledge}`});
        }

        let payload = {
          'messages': messages,
          'model': model,
          'max_tokens': maxToken,
          'temperature': temperature,
          'user': Session.getTemporaryActiveUserKey()
        };

        let functionCalling = false;

        if (browsing) {
          if (messages[messages.length - 1].role !== "function") {
            messages.push({ role: "system", content: "You are able to perform queries on Google search using the function webSearch, then open results and get the content of a web page using the function urlFetch." });
            functions.push(webSearchFunction);
            functions.push(urlFetchFunction);
            let function_calling = { name: "webSearch" };
            payload.function_call = function_calling;
          } else if (messages[messages.length - 1].role == "function" && messages[messages.length - 1].name === "webSearch") {
            let function_calling = { name: "urlFetch" };
            payload.function_call = function_calling;
          }
        }

        if (functions.length >> 0) { // the user has added functions, we enable function calling
          functionCalling = true;
          let payloadFunctions = Object.keys(functions).map(f => ({
            name: functions[f].toJSON().name,
            description: functions[f].toJSON().description,
            parameters: functions[f].toJSON().parameters
          }));
          payload.functions = payloadFunctions;

          if (!payload.function_call) {
            payload.function_call = 'auto';
          }

          if (advancedParametersObject && advancedParametersObject.function_call && JSON.stringify(payload.function_call) !== JSON.stringify({ name: "urlFetch" })) { // the user has set a specific function to call
            let function_calling = { name: advancedParametersObject.function_call };
            payload.function_call = function_calling;
          }
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
              console.log("WARNING : Answer has been troncated because it was too long. To resolve this issue, you can increase the max_tokens property");
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
          return "request failed";
        }

        console.log('Got response from open AI API');
        console.log(responseMessage)

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
                if (typeof functionResponse == "object") {
                  functionResponse = JSON.stringify(functionResponse);
                } else {
                  functionResponse = String(functionResponse);
                }
              }
              if (ENABLE_LOGS) {
                console.log("Conversation stopped because end function has been called");
              }
              return responseMessage;;


            } else if (onlyReturnArguments) {
              if (ENABLE_LOGS) {
                console.log("Conversation stopped because argument return has been enabled - No function has been called");
              }
              return functionArgs;
            } else {
              let functionResponse = callFunction(functionName, functionArgs, argsOrder);
              if (typeof functionResponse != "string") {
                if (typeof functionResponse == "object") {
                  functionResponse = JSON.stringify(functionResponse);
                } else {
                  functionResponse = String(functionResponse);
                }
              }
              if (functionName !== "webSearch" && functionName !== "urlFetch") {
                if (ENABLE_LOGS) {
                  console.log(functionName + "() called by OpenAI.");
                }
              } else if (functionName == "webSearch") {
                payload.function_call = { name: "urlFetch" };
              } else if (functionName == "urlFetch") {
                if (!functionResponse) {
                  if (ENABLE_LOGS) {
                    console.log("The website didn't respond, going back to the results page");
                  }
                  try {
                    let searchResults = JSON.parse(messages[messages.length - 1].content);
                    let newSearchResult = searchResults.filter(function (obj) {
                      return obj.link !== functionArgs.url;
                    });
                    messages[messages.length - 1].content = JSON.stringify(newSearchResult);
                  } catch (e) {
                    console.log(e)
                  }

                  if (advancedParametersObject) {
                    return this.run(advancedParametersObject);
                  } else {
                    return this.run();
                  }
                }
                payload.function_call = "auto";
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
            if (advancedParametersObject) {
              return this.run(advancedParametersObject);
            } else {
              return this.run();
            }

          }
          else {
            // if no function has been found, stop here
            return responseMessage;
          }
        }
        else {
          return responseMessage;
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
      console.log("Function not found or not a function: " + functionName);
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
        console.warn('Error parsing corrected response: ' + e.message);
        return null;
      }
    }
  }

  function webSearch(q) {
    console.log(`Web search : "${q}"`);

    const searchEngineId = "221c662683d054b63";

    let url = `https://www.googleapis.com/customsearch/v1?key=${GoogleCustomSearchAPIKey}&cx=${searchEngineId}&q=${encodeURIComponent(q)}`;

    // If SITE_SEARCH is defined, append site-specific search parameters to the URL
    if (SITE_SEARCH) {
      console.log(`Customed the web search to browse ${SITE_SEARCH}`);
      const site = SITE_SEARCH;
      const siteFilter = 'i';
      url += `&siteSearch=${encodeURIComponent(site)}&siteSearchFilter=${siteFilter}`;
    }

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
    console.log(`Clicked on link : ${url}`);
    const options = {
      'muteHttpExceptions': true
    }
    let response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() == 200) {
      let pageContent = response.getContentText();
      pageContent = sanitizeHtml(pageContent);
      // appsscript.urlFetchWhitelist.filter(item => item !== url);
      return pageContent;

    } else {
      return null;
    }

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

    /**
     * Mandatory
     * @param {string} apiKey - Your Open AI API key.
     */
    setOpenAIAPIKey: function (apiKey) {
      OpenAIKey = apiKey;
    },

    /**
     * If you want to enable browsing
     * @param {string} apiKey - Your Google API key.
     */
    setGoogleSearchAPIKey: function (apiKey) {
      GoogleCustomSearchAPIKey = apiKey;
    },

    /**
     * If you want to limit your web browsing to one web page
     * @param {string} url -  the url of the website you want to browse
     */
    restrictToWebsite: function (url) {
      SITE_SEARCH = url;
    },

    /**
     * If you want to add the content of a web page to the chat
     * @param {string} url - the url of the webpage you want to fetch
     */
    addKnowledgeLink: function (url) {
      KNOWLEDGE_LINK = url;
    },

    /**
     * If you only want to keep your own logs and disable those of the library
     */
    disableLogs: function () {
      ENABLE_LOGS = false;
    }
  }
}
)();