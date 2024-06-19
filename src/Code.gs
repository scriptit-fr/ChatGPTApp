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

  let openAIKey = "";

  let googleCustomSearchAPIKey = "";
  let restrictSearch;

  let verbose = true;

  const noResultFromWebSearchMessage = `Your search did not match any documents. 
  Try with different, more general or fewer keywords.`;

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
      this.onlyReturnArguments = function (bool) {
        if (bool) {
          onlyArgs = true;
        }
        return this;
      }

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
    .setDescription("Perform a web search via the Google Custom Search JSON API. Returns an array of search results (including the URL, title and plain text snippet for each result)")
    .addParameter("q", "string", "the query for the web search.");

  let urlFetchFunction = new FunctionObject()
    .setName("urlFetch")
    .setDescription("Fetch the viewable content of a web page. HTML tags will be stripped, returning a text-only version.")
    .addParameter("url", "string", "The URL to fetch.");

  let imageDescriptionFunction = new FunctionObject()
    .setName("getImageDescription")
    .setDescription("To retrieve the description of an image.")
    .addParameter("imageUrl", "string", "The URL of the image.")
    .addParameter("highFidelity", "boolean", `Default: false. To improve the image quality, not needed in most cases.`, isOptional = true);

  /**
   * @class
   * Class representing a chat.
   */
  class Chat {
    constructor() {
      let messages = [];
      let tools = [];
      let model = "gpt-3.5-turbo"; // default 
      let temperature = 0.5;
      let max_tokens = 300;
      let browsing = false;
      let vision = false;
      let onlyRetrieveSearchResults = false;
      let knowledgeLink;

      let webSearchQueries = [];
      let webPagesOpened = [];

      let maximumAPICalls = 30;
      let numberOfAPICalls = 0;

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
      this.addFunction = function (functionObject) {
        tools.push({
          type: "function",
          function: functionObject
        });
        return this;
      };

      /**
      * Add an image to the chat. Will automatically include an image with gpt-4-turbo-2024-04-09.
      * @param {string} imageUrl - The URL of the image to add.
      * @returns {Chat} - The current Chat instance.
      */
      this.addImage = function (imageUrl) {
        messages.push(
          {
            role: "user",
            content: [
              {
                type: "image_url",
                image_url: { url: imageUrl }
              }
            ]
          }
        );
        vision = true;
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
       * Get the tools of the chat.
       * returns {FunctionObject[]} - The tools of the chat.
       */
      this.getFunctions = function () {
        return JSON.stringify(tools);
      };

      /**
       * Disable logs generated by this library
       * @returns {Chat} - The current Chat instance.
       */
      this.disableLogs = function (bool) {
        if (bool) {
          verbose = false;
        }
        return this;
      };

      /**
       * OPTIONAL
       * 
       * Allow openAI to browse the web.
       * @param {true|"only_retrieve_results"} scope - set to true to enable full browsing, or "only_retrieve_results" to search Google without opening web pages. 
       * @param {string} [urlOrsearchEngineId] - A specific site you want to restrict the search on or a Search engine ID. 
       * @returns {Chat} - The current Chat instance.
       */
      this.enableBrowsing = function (scope, urlOrsearchEngineId) {
        if (scope) {
          browsing = true;
          if (scope == "only_retrieve_results") {
            onlyRetrieveSearchResults = true;
          }
        }
        if (urlOrsearchEngineId) {
          restrictSearch = urlOrsearchEngineId;
        }
        return this;
      };

      /**
       * OPTIONAL
       * 
       * Allow openAI to call vision model.
       * @param {true} scope - set to true to enable vision. 
       * @returns {Chat} - The current Chat instance.
       */
      this.enableVision = function (scope) {
        if (scope) {
          vision = true;
        }
        return this;
      };

      /**
       * Includes the content of a web page in the prompt sent to openAI
       * @param {string} url - the url of the webpage you want to fetch
       * @returns {Chat} - The current Chat instance.
       */
      this.addKnowledgeLink = function (url) {
        knowledgeLink = url;
        return this;
      };

      /**
       * If you want to limit the number of calls to the OpenAI API
       * A good way to avoid infinity loops and manage your budget.
       * @param {number} maxAPICalls - 
       */
      this.setMaximumAPICalls = function (maxAPICalls) {
        maximumAPICalls = maxAPICalls;
      };

      this.toJson = function () {
        return {
          messages: messages,
          tools: tools,
          model: model,
          temperature: temperature,
          max_tokens: max_tokens,
          browsing: browsing,
          webSearchQueries: webSearchQueries,
          webPagesOpened: webPagesOpened,
          maximumAPICalls: maximumAPICalls,
          numberOfAPICalls: numberOfAPICalls
        };
      };

      /**
       * Start the chat conversation.
       * Sends all your messages and eventual function to chat GPT.
       * Will return the last chat answer.
       * If a function calling model is used, will call several functions until the chat decides that nothing is left to do.
       * @param {Object} [advancedParametersObject] OPTIONAL - For more advanced settings and specific usage only. {model, temperature, function_call}
       * @param {"gpt-3.5-turbo" | "gpt-3.5-turbo-16k" | "gpt-4" | "gpt-4-32k" | "gpt-4-1106-preview" | "gpt-4-turbo-preview" | "gpt-4o"} [advancedParametersObject.model]
       * @param {number} [advancedParametersObject.temperature]
       * @param {number} [advancedParametersObject.max_tokens]
       * @param {string} [advancedParametersObject.function_call]
       * @returns {object} - the last message of the chat 
       */
      this.run = function (advancedParametersObject) {
        if (!openAIKey) {
          if (googleCustomSearchAPIKey) {
            throw Error("Careful to use setOpenAIAPIKey to set your OpenAI API key and not setGoogleSearchAPIKey.");
          }
          else {
            throw Error("Please set your OpenAI API key using ChatGPTApp.setOpenAIAPIKey(youAPIKey)");
          }
        }
        if (browsing && !googleCustomSearchAPIKey) {
          throw Error("Please set your Google custom search API key using ChatGPTApp.setGoogleSearchAPIKey(youAPIKey)");
        }

        if (advancedParametersObject) {
          if (advancedParametersObject.model) {
            model = advancedParametersObject.model;
          }
          if (advancedParametersObject.temperature) {
            temperature = advancedParametersObject.temperature;
          }
          if (advancedParametersObject.max_tokens) {
            max_tokens = advancedParametersObject.max_tokens;
          }
        }

        if (knowledgeLink) {
          let knowledge = urlFetch(knowledgeLink);
          if (!knowledge) {
            throw Error(`The webpage ${knowledgeLink} didn't respond, please change the url of the addKnowledgeLink() function.`);
          }
          messages.push({
            role: "system",
            content: `Information to help with your response (publicly available here: ${knowledgeLink}):\n\n${knowledge}`
          });
          knowledgeLink = null;
        }

        let payload = {
          'messages': messages,
          'model': model,
          'max_tokens': max_tokens,
          'temperature': temperature,
          'user': Session.getTemporaryActiveUserKey()
        };

        let functionCalling = false;

        if (browsing) {
          if (messages[messages.length - 1].role !== "tool") {
            tools.push({
              type: "function",
              function: webSearchFunction
            });
            let messageContent = `You are able to perform search queries on Google using the function webSearch and read the search results. `;
            if (!onlyRetrieveSearchResults) {
              messageContent += "Then you can select a search result and read the page content using the function urlFetch. ";
              tools.push({
                type: "function",
                function: urlFetchFunction
              });
            }
            messages.push({
              role: "system",
              content: messageContent
            });
            payload.tool_choice = {
              type: "function",
              function: { name: "webSearch" }
            };
          }
          else if (messages[messages.length - 1].role == "tool" &&
            messages[messages.length - 1].name === "webSearch" &&
            !onlyRetrieveSearchResults) {
            if (messages[messages.length - 1].content !== noResultFromWebSearchMessage) {
              // force openAI to call the function urlFetch after retrieving results for a particular search
              payload.tool_choice = {
                type: "function",
                function: { name: "urlFetch" }
              };
            }
          }
        }

        if (vision) {
          payload.model = 'gpt-4o';
        }

        if (tools.length >> 0) {
          // the user has added functions, enable function calling
          functionCalling = true;
          let payloadTools = Object.keys(tools).map(t => ({
            type: "function",
            function: {
              name: tools[t].function.toJSON().name,
              description: tools[t].function.toJSON().description,
              parameters: tools[t].function.toJSON().parameters
            }
          }));
          payload.tools = payloadTools;

          if (!payload.tool_choice) {
            payload.tool_choice = 'auto';
          }

          if (advancedParametersObject?.function_call &&
            payload.tool_choice.function?.name !== "urlFetch" &&
            payload.tool_choice.function?.name !== "webSearch") {
            // the user has set a specific function to call
            let tool_choosing = {
              type: "function",
              function: {
                name: advancedParametersObject.function_call
              }
            };
            payload.tool_choice = tool_choosing;
          } else if (messages[messages.length - 1].role == "tool" && messages[messages.length - 1].name == urlFetch) {
            // Once we've opened a web page,
            // let the model decide what to do
            // eg: model can either be satisfied with the info found in the web page or decide to open another web page
            payload.tool_choice = 'auto'
          }
        }
        let responseMessage;
        if (numberOfAPICalls <= maximumAPICalls) {
          responseMessage = callOpenAIApi(payload);
          numberOfAPICalls++;
        } else {
          throw new Error(`Too many calls to OpenAI API: ${numberOfAPICalls}`);
        }

        if (functionCalling) {
          // Check if GPT wanted to call a function
          if (responseMessage.tool_calls) {
            messages = handleToolCalls(responseMessage, tools, messages, webSearchQueries, webPagesOpened);
            // check if endWithResults or onlyReturnArguments
            if (messages[messages.length - 1].role == "system") {
              if (messages[messages.length - 1].content == "endWithResult") {
                if (verbose) {
                  console.log("Conversation stopped because end function has been called");
                }
                return messages[messages.length - 2]; // the last chat completion
              } else if (messages[messages.length - 1].content == "onlyReturnArguments") {
                if (verbose) {
                  console.log("Conversation stopped because argument return has been enabled - No function has been called");
                }
                return parseResponse(messages[messages.length - 2].tool_calls[0].function.arguments); // the argument(s) of the last function called
              }
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

  function callOpenAIApi(payload) {
    let maxRetries = 5;
    let retries = 0;
    let success = false;

    let responseMessage, finish_reason;
    while (retries < maxRetries && !success) {
      let options = {
        'method': 'post',
        'headers': {
          'Content-Type': 'application/json',
          'Authorization': 'Bearer ' + openAIKey
        },
        'payload': JSON.stringify(payload),
        'muteHttpExceptions': true
      };

      let response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', options);
      let responseCode = response.getResponseCode();

      if (responseCode === 200) {
        // The request was successful, exit the loop.
        const parsedResponse = JSON.parse(response.getContentText());
        responseMessage = parsedResponse.choices[0].message;
        finish_reason = parsedResponse.choices[0].finish_reason;
        if (finish_reason == "length") {
          console.warn(`OpenAI response has been troncated because it was too long. To resolve this issue, you can increase the max_tokens property. max_tokens: ${payload.max_tokens}, prompt_tokens: ${parsedResponse.usage.prompt_tokens}, completion_tokens: ${parsedResponse.usage.completion_tokens}`);
        }
        success = true;
      }
      else if (responseCode === 429) {
        console.warn(`Rate limit reached when calling openAI API, will automatically retry in a few seconds.`);
        // Rate limit reached, wait before retrying.
        let delay = Math.pow(2, retries) * 1000; // Delay in milliseconds, starting at 1 second.
        Utilities.sleep(delay);
        retries++;
      }
      else if (responseCode === 503 || responseCode === 500) {
        // The server is temporarily unavailable, or an issue occured on OpenAI servers. wait before retrying.
        // https://platform.openai.com/docs/guides/error-codes/api-errors
        let delay = Math.pow(2, retries) * 1000; // Delay in milliseconds, starting at 1 second.
        Utilities.sleep(delay);
        retries++;
      }
      else {
        // The request failed for another reason, log the error and exit the loop.
        console.error(`Request to openAI failed with response code ${responseCode} - ${response.getContentText()}`);
        break;
      }
    }

    if (!success) {
      throw new Error(`Failed to call openAI API after ${retries} retries.`);
    }

    if (verbose) {
      console.log(`Got response from openAI API.`);
    }

    return responseMessage;
  }


  /**
   * Handles the invocation of tool functions based on the response from a chat session
   * 
   * First inform the chat that the function has been called
   * https://platform.openai.com/docs/guides/function-calling
   * Chat needs it's own response to get the tool_calls id(s), and won't go on if the following messages are not the correct tool_calls
   * 
   * Then iterates over the tool calls in the `responseMessage` to find matching functions in the `tools`
   * array. It then calls these functions with the parsed arguments. Depending on the function's specification
   * (e.g., if it's an ending function or only returns arguments), it may alter the flow of the conversation
   * by updating the `messages` array with new messages or by returning early.
   * 
   * @param {Object} responseMessage - The response message object from the chat that contains the tool calls.
   * @param {Object[]} tools - An array of tool objects that describe available tools and their specifications
   * @param {Object[]} messages - The array of message objects to be updated with responses from tool calls.
   * @param {string[]} webSearchQueries - An array to record search queries made as a result of the tool calls.
   * @param {string[]} webPagesOpened - An array to record web pages opened as a result of the tool calls.
   * @returns {Object[]} The updated array of message objects after processing the tool calls.
   */
  function handleToolCalls(responseMessage, tools, messages, webSearchQueries, webPagesOpened) {
    messages.push(responseMessage);
    for (let tool_call in responseMessage.tool_calls) {
      if (responseMessage.tool_calls[tool_call].type == "function") {
        // Call the function
        let functionName = responseMessage.tool_calls[tool_call].function.name;
        let functionArgs = parseResponse(responseMessage.tool_calls[tool_call].function.arguments);

        let argsOrder = [];
        let endWithResult = false;
        let onlyReturnArguments = false;

        for (let t in tools) {
          let currentFunction = tools[t].function.toJSON();
          if (currentFunction.name == functionName) {
            argsOrder = currentFunction.argumentsInRightOrder; // get the args in the right order
            endWithResult = currentFunction.endingFunction;
            onlyReturnArguments = currentFunction.onlyArgs;
            break;
          }
        }

        if (endWithResult) {
          // User defined that if this function has been called, then no more actions should be performed with the chat.
          let functionResponse = callFunction(functionName, functionArgs, argsOrder);
          if (typeof functionResponse != "string") {
            if (typeof functionResponse == "object") {
              functionResponse = JSON.stringify(functionResponse);
            }
            else {
              functionResponse = String(functionResponse);
            }
          }
          messages.push({
            "role": "system",
            "content": "endWithResult"
          });
          return messages;
        }
        else if (onlyReturnArguments) {
          messages.push({
            "role": "system",
            "content": "onlyReturnArguments"
          });
          return messages;
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

          if (functionName == "webSearch") {
            webSearchQueries.push(functionArgs.q);
          }
          else if (functionName == "urlFetch") {
            webPagesOpened.push(functionArgs.url);
            if (!functionResponse) {
              if (verbose) {
                console.log("The website didn't respond, going back to search results.");
              }
              let searchResults = JSON.parse(messages[messages.length - 1].content);
              let updatedSearchResults = searchResults.filter(function (obj) {
                return obj.link !== functionArgs.url;
              });
              messages[messages.length - 1].content = JSON.stringify(updatedSearchResults);
            }
            if (verbose) {
              console.log("Web page opened, let model decide what to do next (open another web page or perform another action).");
            }
          }
          else {
            if (verbose) {
              console.log(`function ${functionName}() called by OpenAI.`);
            }
          }
          messages.push({
            "tool_call_id": responseMessage.tool_calls[tool_call].id,
            "role": "tool",
            "name": functionName,
            "content": functionResponse,
          });
        }
      }
    }
    return messages;
  }

  function callFunction(functionName, jsonArgs, argsOrder) {
    // Handle internal functions
    if (functionName == "webSearch") {
      return webSearch(jsonArgs.q);
    }
    if (functionName == "urlFetch") {
      return urlFetch(jsonArgs.url);
    }
    if (functionName == "getImageDescription") {
      if (jsonArgs.fidelity) {
        return getImageDescription(jsonArgs.imageUrl, jsonArgs.fidelity);
      } else {
        return getImageDescription(jsonArgs.imageUrl);
      }
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
        return "The function has been sucessfully executed but has nothing to return";
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
    // https://programmablesearchengine.google.com/controlpanel/overview?cx=221c662683d054b63
    let searchEngineId = "221c662683d054b63";
    let url = `https://www.googleapis.com/customsearch/v1?key=${googleCustomSearchAPIKey}`;

    // If restrictSearch is defined, check wether to restrict to a specific site or use a specific Search Engine
    if (restrictSearch) {
      if (restrictSearch.includes('.')) {
        // Search restricted to specific site
        if (verbose) {
          console.log(`Site search on ${restrictSearch}`);
        }
        url += `&siteSearch=${encodeURIComponent(restrictSearch)}&siteSearchFilter=i`;
      }
      else {
        // Use the desired Search Engine
        // https://programmablesearchengine.google.com/controlpanel/all
        searchEngineId = restrictSearch;
      }
    }
    url += `&cx=${searchEngineId}&q=${encodeURIComponent(q)}&num=10`;

    const urlfetchResp = UrlFetchApp.fetch(url);
    const resp = JSON.parse(urlfetchResp.getContentText());

    let searchResults;
    if (!resp.items?.length) {
      searchResults = noResultFromWebSearchMessage;
    }
    else {
      // https://developers.google.com/custom-search/v1/reference/rest/v1/Search?hl=en#Result
      searchResults = JSON.stringify(resp.items.slice(0, 10).map(function (item) {
        return {
          title: item.title,
          link: item.link,
          snippet: item.snippet
        };
        // filter to remove undefined values from the results array
      }).filter(Boolean));
    }

    const nbOfResults = resp.searchInformation.totalResults;
    if (verbose) {
      Logger.log({
        message: `Web search : "${q}" - ${nbOfResults} results`,
        searchResults: searchResults
      });
    }

    return searchResults;
  }


  function urlFetch(url) {
    if (verbose) {
      console.log(`Clicked on link : ${url}`);
    }
    let response;
    try {
      response = UrlFetchApp.fetch(url);
    }
    catch (e) {

      console.warn(`Error fetching the URL: ${e.message}`);
      return JSON.stringify({
        error: "Failed to fetch the URL : You are not authorized to access this website. Try another one."
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

  function getImageDescription(imageUrl, fidelity) {
    const extensions = ['png', 'jpeg', 'gif', 'webp'];
    if (!extensions.some(extension => imageUrl.toLowerCase().endsWith(`.${extension}`))) {
      if (verbose) {
        Logger.log({
          message: `Tried to get description of image, but extension is not supported by OpenAI API. Supported extensions are 'png', 'jpeg', 'gif', 'webp'`,
          imageUrl: imageUrl
        })
      }
      return "This is not a supported image, no description available."
    }

    if (!fidelity) {
      fidelity = "low";
    } else {
      fidelity = "high";
    }

    let imageMessage = [{
      role: "user",
      content: [
        {
          type: "text",
          text: "What is the content of this image ? Focus on important element." // Give a more human friendly description than the initial example prompt given by OpenAI
        },
        {
          type: "image_url",
          image_url: {
            url: imageUrl,
            detail: fidelity
          }
        },
      ]
    }];

    let payload = {
      'messages': imageMessage,
      'model': "gpt-4-vision-preview",
      'max_tokens': 1000,
      'user': Session.getTemporaryActiveUserKey()
    };

    let responseMessage = callOpenAIApi(payload);

    return responseMessage;
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

    // Images - Updated to use Markdown syntax
    htmlString = htmlString.replace(/<img src="(.+?)" alt="(.*?)" ?\/?>/g, '![$2]($1)'); // Markdown syntax for images is ![alt text](image URL).

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
     * @param {string} apiKey - Your openAI API key.
     */
    setOpenAIAPIKey: function (apiKey) {
      openAIKey = apiKey;
    },

    /**
     * If you want to enable browsing
     * @param {string} apiKey - Your Google API key.
     */
    setGoogleSearchAPIKey: function (apiKey) {
      googleCustomSearchAPIKey = apiKey;
    },

    /**
     * If you want to acces what occured during the conversation
     * @param {Chat} chat - your chat instance.
     * @returns {object} - the web search queries, the web pages opened and an historic of all the messages of the chat
     */
    debug: function (chat) {
      return {
        getWebSearchQueries: function () {
          return chat.toJson().webSearchQueries
        },
        getWebPagesOpened: function () {
          return chat.toJson().webPagesOpened
        }
      }
    }
  }
})();
