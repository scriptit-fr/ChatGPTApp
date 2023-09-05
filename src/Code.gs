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
const ChatGPTApp = (function() {

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
      this.setName = function(nameOfYourFunction) {
        name = nameOfYourFunction;
        return this;
      };

      /**
       * Sets the description of a function.
       * @param {string} descriptionOfYourFunction - The description to set for the function.
       * @returns {FunctionObject} - The current Function instance.
       */
      this.setDescription = function(descriptionOfYourFunction) {
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
      this.endWithResult = function(bool) {
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
      this.addParameter = function(name, type, description, isOptional = false) {
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
          }
          else {
            throw Error("Please precise the type of the items contained in the array when calling addParameter. Use format Array.<itemsType> for the type parameter.");
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
      this.onlyReturnArguments = function(bool) {
        if (bool) {
          onlyArgs = true;
        }
        return this;
      }

      this.toJSON = function() {
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

      let webSearchQueries = [];
      let webPagesOpened = [];

      /**
       * Add a message to the chat.
       * @param {string} messageContent - The message to be added.
       * @param {boolean} [system] - OPTIONAL - True if message from system, False for user. 
       * @returns {Chat} - The current Chat instance.
       */
      this.addMessage = function(messageContent, system) {
        let role = "user";
        if (system) {
          role = "system";
        }
        messages.push({
          role: role,
          content: messageContent
        });
        return this;
      };

      /**
       * Add a function to the chat.
       * @param {FunctionObject} functionObject - The function to be added.
       * @returns {Chat} - The current Chat instance.
       */
      this.addFunction = function(functionObject) {
        functions.push(functionObject);
        return this;
      };

      /**
       * Get the messages of the chat.
       * returns {string[]} - The messages of the chat.
       */
      this.getMessages = function() {
        return JSON.stringify(messages);
      };

      /**
       * Get the functions of the chat.
       * returns {FunctionObject[]} - The functions of the chat.
       */
      this.getFunctions = function() {
        return JSON.stringify(functions);
      };

      /**
       * If you only want to keep your own logs and disable those of the library
       * @returns {Chat} - The current Chat instance.
       */
      this.disableLogs = function(bool) {
          if (bool) {
            ENABLE_LOGS = false;
          }
          return this;
        },

        /**
         * OPTIONAL
         * 
         * Enable the chat to use a google serach engine to browse the web.
         * @param {boolean} bool - Whether or not you wish for the option to be enabled. 
         * @param {string} [url] - A specific web page url you want to restrict the search search to. 
         * @returns {Chat} - The current Chat instance.
         */
        this.enableBrowsing = function(bool, url) {
          if (bool) {
            browsing = true
          }
          if (url) {
            SITE_SEARCH = url;
          }
          return this;
        };

      /**
       * If you want to add the content of a web page to the chat
       * @param {string} url - the url of the webpage you want to fetch
       * @returns {Chat} - The current Chat instance.
       */
      this.addKnowledgeLink = function(url) {
        KNOWLEDGE_LINK = url;
        return this;
      };

      this.toJson = function() {
        return {
          messages: messages,
          functions: functions,
          model: model,
          temperature: temperature,
          maxToken: maxToken,
          browsing: browsing,
          webSearchQueries: webSearchQueries,
          webPagesOpened: webPagesOpened
        };
      };

      /**
       * Start the chat conversation.
       * Sends all your messages and eventual function to chat GPT.
       * Will return the last chat answer.
       * If a function calling model is used, will call several functions until the chat decides that nothing is left to do.
       * @param {{model?: "gpt-3.5-turbo" | "gpt-3.5-turbo-16k" | "gpt-4" | "gpt-4-32k", temperature?: number, maxToken?: number, function_call?: string}} [advancedParametersObject] - OPTIONAL - For more advanced settings and specific usage only. {model, temperature, function_call}
       * @returns {object} - the last message of the chat 
       */
      this.run = function(advancedParametersObject) {
        if (!OpenAIKey) {
          if (GoogleCustomSearchAPIKey) {
            throw Error("Careful to use setOpenAIAPIKey to set your OpenAI API key and not setGoogleSearchAPIKey.");
          }
          else {
            throw Error("Please set your OpenAI API key using ChatGPTApp.setOpenAIAPIKey(youAPIKey)");
          }
        }
        if (browsing && !GoogleCustomSearchAPIKey) {
          throw Error("Please set your Google custom search API key using ChatGPTApp.setGoogleSearchAPIKey(youAPIKey)");
        }

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
          if (!knowledge) {
            throw Error(`The webpage ${KNOWLEDGE_LINK} didn't respond, please change the url of the addKnowledgeLink() function.`)
          }
          messages.push({
            role: "system",
            content: `Here's an article to help:\n\n${knowledge}`
          });
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
            messages.push({
              role: "system",
              content: "You are able to perform queries on Google search using the function webSearch, then open results and get the content of a web page using the function urlFetch."
            });
            functions.push(webSearchFunction);
            functions.push(urlFetchFunction);
            let function_calling = {
              name: "webSearch"
            };
            payload.function_call = function_calling;
          }
          else if (messages[messages.length - 1].role == "function" &&
            messages[messages.length - 1].name === "webSearch") {
            let function_calling = {
              name: "urlFetch"
            };
            payload.function_call = function_calling;
          }
        }

        if (functions.length >> 0) {
          // the user has added functions, enable function calling
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

          if (advancedParametersObject &&
            advancedParametersObject.function_call &&
            JSON.stringify(payload.function_call) !== JSON.stringify({
              name: "urlFetch"
            })) {
            // the user has set a specific function to call
            let function_calling = {
              name: advancedParametersObject.function_call
            };
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
              console.warn("OpenAI response has been troncated because it was too long. To resolve this issue, you can increase the max_tokens property");
            }
            success = true;
          }
          else if (responseCode === 503) {
            // The server is temporarily unavailable, wait before retrying.
            let delay = Math.pow(2, attempt) * 1000; // Delay in milliseconds, starting at 1 second.
            Utilities.sleep(delay);
            attempt++;
          }
          else {
            // The request failed for another reason, log the error and exit the loop.
            console.error('Request failed with response code', responseCode);
            break;
          }
        }

        if (!success) {
          console.error('Failed to call openAI API after', maxAttempts, 'attempts');
          return "request failed";
        }

        if (ENABLE_LOGS) {
          console.log('Got response from openAI API');
        }

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
                }
                else {
                  functionResponse = String(functionResponse);
                }
              }
              if (ENABLE_LOGS) {
                console.log("Conversation stopped because end function has been called");
              }
              return responseMessage;;


            }
            else if (onlyReturnArguments) {
              if (ENABLE_LOGS) {
                console.log("Conversation stopped because argument return has been enabled - No function has been called");
              }
              return functionArgs;
            }
            else {
              let functionResponse = callFunction(functionName, functionArgs, argsOrder);
              if (typeof functionResponse != "string") {
                if (typeof functionResponse == "object") {
                  functionResponse = JSON.stringify(functionResponse);
                }
                else {
                  functionResponse = String(functionResponse);
                }
              }
              if (functionName !== "webSearch" && functionName !== "urlFetch") {
                if (ENABLE_LOGS) {
                  console.log(functionName + "() called by OpenAI.");
                }
              }
              else if (functionName == "webSearch") {
                webSearchQueries.push(functionArgs.q);
                payload.function_call = {
                  name: "urlFetch"
                };
              }
              else if (functionName == "urlFetch") {
                webPagesOpened.push(functionArgs.url);
                if (!functionResponse) {
                  if (ENABLE_LOGS) {
                    console.log("The website didn't respond, going back to the results page");
                  }
                  try {
                    let searchResults = JSON.parse(messages[messages.length - 1].content);
                    let newSearchResult = searchResults.filter(function(obj) {
                      return obj.link !== functionArgs.url;
                    });
                    messages[messages.length - 1].content = JSON.stringify(newSearchResult);
                  }
                  catch (e) {
                    console.log(e);
                  }

                  if (advancedParametersObject) {
                    return this.run(advancedParametersObject);
                  }
                  else {
                    return this.run();
                  }
                }
                payload.function_call = "auto";
              }
              // Inform the chat that the function has been called
              messages.push({
                "role": "assistant",
                "content": null,
                "function_call": {
                  "name": functionName,
                  "arguments": JSON.stringify(functionArgs)
                }
              });
              messages.push({
                "role": "function",
                "name": functionName,
                "content": functionResponse,
              })
            }
            if (advancedParametersObject) {
              return this.run(advancedParametersObject);
            }
            else {
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
      }
      else {
        return "the function has been sucessfully executed but has nothing to return";
      }
    }
    else {
      throw Error("Function not found or not a function: " + functionName);
    }
  }

  function parseResponse(response) {
    try {
      let parsedReponse = JSON.parse(response);
      return parsedReponse;
    }
    catch (e) {
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
        line = line.trimEnd(); // Strip trailing white spaces
        if (line[line.length - 1] !== ',') {
          if (line[line.length - 1] == ':') {
            // If line has the format "property":, add null,
            lines[i] = line + ' ""';
          }
          else if (!line.includes(':')) {
            lines[i] = line + ': ""';
          }
          else if (line[line.length - 1] !== '"') {
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
      }
      catch (e) {
        // If parsing still fails, log the error and return null.
        console.warn('Error parsing corrected response: ' + e.message);
        return null;
      }
    }
  }

  function webSearch(q) {
    if (ENABLE_LOGS) {
      console.log(`Web search : "${q}"`);
    }

    const searchEngineId = "221c662683d054b63";

    let url = `https://www.googleapis.com/customsearch/v1?key=${GoogleCustomSearchAPIKey}&cx=${searchEngineId}&q=${encodeURIComponent(q)}`;

    // If SITE_SEARCH is defined, append site-specific search parameters to the URL
    if (SITE_SEARCH) {
      if (ENABLE_LOGS) {
        console.log(`Site search on ${SITE_SEARCH}`);
      }
      const site = SITE_SEARCH;
      const siteFilter = 'i';
      url += `&siteSearch=${encodeURIComponent(site)}&siteSearchFilter=${siteFilter}`;
    }

    let response = UrlFetchApp.fetch(url);
    let data = JSON.parse(response.getContentText());

    let resultsInfo = [];

    if (data.items) {
      resultsInfo = data.items.map(function(item) {
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
    if (ENABLE_LOGS) {
      console.log(`Clicked on link : ${url}`);
    }
    const options = {
      'muteHttpExceptions': true
    }
    let response;
    try {
      response = UrlFetchApp.fetch(url);
    }
    catch (e) {
      console.error(`Error fetching the URL: ${e.message}`);
      return JSON.stringify({
        error: "Failed to fetch the URL"
      });
    }
    if (response.getResponseCode() == 200) {
      let pageContent = response.getContentText();
      pageContent = convertHtmlToMarkdown(pageContent);
      return pageContent;
    }
    else {
      return null;
    }

  }

  function convertHtmlToMarkdown(htmlString) {
    // Remove <script> tags and their content
    htmlString = htmlString.replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '');
    // Remove <style> tags and their content
    htmlString = htmlString.replace(/<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>/gi, '');
    // Remove on* attributes (e.g., onclick, onload)
    htmlString = htmlString.replace(/ ?on[a-z]*=".*?"/gi, '');
    // Remove style attributes
    htmlString = htmlString.replace(/ ?style=".*?"/gi, '');
    // Remove class attributes
    htmlString = htmlString.replace(/ ?class=".*?"/gi, '');

    // Convert &nbsp; to spaces
    htmlString = htmlString.replace(/&nbsp;/g, ' ');

    // Improved anchor tag conversion
    htmlString = htmlString.replace(/<a [^>]*href="(.*?)"[^>]*>(.*?)<\/a>/g, '[$2]($1)');

    // Strong
    htmlString = htmlString.replace(/<strong>(.*?)<\/strong>/g, '**$1**');
    // Emphasize
    htmlString = htmlString.replace(/<em>(.*?)<\/em>/g, '_$1_');
    // Headers
    for (let i = 1; i <= 6; i++) {
      const regex = new RegExp(`<h${i}>(.*?)<\/h${i}>`, 'g');
      htmlString = htmlString.replace(regex, `${'#'.repeat(i)} $1`);
    }

    // Blockquote
    htmlString = htmlString.replace(/<blockquote>(.*?)<\/blockquote>/g, '> $1');

    // Unordered list
    htmlString = htmlString.replace(/<ul>(.*?)<\/ul>/g, '$1');
    htmlString = htmlString.replace(/<li>(.*?)<\/li>/g, '- $1');

    // Ordered list
    htmlString = htmlString.replace(/<ol>(.*?)<\/ol>/g, '$1');
    htmlString = htmlString.replace(/<li>(.*?)<\/li>/g, (match, p1, offset, string) => {
      const number = string.substring(0, offset).match(/<li>/g).length;
      return `${number}. ${p1}`;
    });

    // Handle table headers (if they exist)
    htmlString = htmlString.replace(/<thead>([\s\S]*?)<\/thead>/g, (match, content) => {
      let headerRow = content.replace(/<th>(.*?)<\/th>/g, '| $1 ').trim() + '|';
      let separatorRow = headerRow.replace(/[^|]+/g, match => {
        return '-'.repeat(match.length);
      }).replace(/\| -/g, '|--');
      return headerRow + '\n' + separatorRow;
    });

    // Handle table rows
    htmlString = htmlString.replace(/<tbody>([\s\S]*?)<\/tbody>/g, (match, content) => {
      return content.replace(/<tr>([\s\S]*?)<\/tr>/g, (match, trContent) => {
        return trContent.replace(/<td>(.*?)<\/td>/g, '| $1 ').trim() + '|';
      });
    });

    // Inline code
    htmlString = htmlString.replace(/<code>(.*?)<\/code>/g, '`$1`');
    // Preformatted text
    htmlString = htmlString.replace(/<pre>(.*?)<\/pre>/g, '```\n$1\n```');
    // Images with URLs as sources
    htmlString = htmlString.replace(/<img src="(.+?)" alt="(.*?)" ?\/?>/g, '[Image here]');

    // Remove remaining HTML tags
    htmlString = htmlString.replace(/<[^>]*>/g, '');

    // Trim excessive white spaces between words/phrases
    htmlString = htmlString.replace(/ +/g, ' ');

    // Remove whitespace followed by newline patterns
    htmlString = htmlString.replace(/ \n/g, '\n');

    // Normalize the line endings to just \n
    htmlString = htmlString.replace(/\r\n/g, '\n');

    // Collapse multiple contiguous newline characters down to a single newline
    htmlString = htmlString.replace(/\n{2,}/g, '\n');

    // Trim leading and trailing white spaces and newlines
    htmlString = htmlString.trim();

    return htmlString;
  }

  return {
    /**
     * Create a new chat.
     * @params {string} apiKey - Your openAI API key.
     * @returns {Chat} - A new Chat instance.
     */
    newChat: function() {
      return new Chat();
    },

    /**
     * Create a new function.
     * @returns {FunctionObject} - A new Function instance.
     */
    newFunction: function() {
      return new FunctionObject();
    },

    /**
     * Mandatory
     * @param {string} apiKey - Your openAI API key.
     */
    setOpenAIAPIKey: function(apiKey) {
      OpenAIKey = apiKey;
    },

    /**
     * If you want to enable browsing
     * @param {string} apiKey - Your Google API key.
     */
    setGoogleSearchAPIKey: function(apiKey) {
      GoogleCustomSearchAPIKey = apiKey;
    },

    /**
     * If you want to acces what occured during the conversation
     * @param {Chat} chat - your chat instance.
     * @returns {object} - the web search queries, the web pages opened and an historic of all the messages of the chat
     */
    debug: function(chat) {
      return {
        getWebSearchQueries: function() {
          return chat.toJson().webSearchQueries
        },
        getWebPagesOpened: function() {
          return chat.toJson().webPagesOpened
        }
      }
    }
  }
})();
