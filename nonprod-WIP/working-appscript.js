
const scriptProperties = PropertiesService.getScriptProperties();//;
const GEMINI_API_KEY = scriptProperties.getProperty("ai_key"); 

// For better security, consider storing the API key in Properties Service:
// https://developers.google.com/apps-script/guides/properties

function onOpen() {
  SlidesApp.getUi()
      .createMenu('Create a shell HLD')
      .addItem('Create from Prompt...', 'startGenerationProcess')
      .addToUi();
}


function extractPresentationId(url) {
    if (!url) return null;
    const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
    return match ? match[1] : null;
}



function startGenerationProcess() {
  const ui = SlidesApp.getUi();

  // This has been edited out to decrease token input and only have presentation + prompt
  // 1. Prompt for the example presentation URL first.
  // const exampleUrlResponse = ui.prompt(
  //   'Step 1: Provide Example Presentation',
  //   'Please enter the full URL of the example Google Slides presentation:',
  //   ui.ButtonSet.OK_CANCEL
  // );

  // Exit if the user cancels or provides no input
  // if (exampleUrlResponse.getSelectedButton() !== ui.Button.OK || !exampleUrlResponse.getResponseText()) {
  //   return;
  // }
  const exampleUrl = ''//exampleUrlResponse.getResponseText();
  const examplePresentationId = extractPresentationId(exampleUrl);

  // Validate the extracted ID
  // if (!examplePresentationId) {
  //     ui.alert('Invalid URL', 'The URL provided does not appear to be a valid Google Slides presentation URL. Please try again.', ui.ButtonSet.OK);
  //     return;
  // }

  // 2. Prompt the user for the core information to generate the presentation.
  const promptResponse = ui.prompt(
    'Step 2: Describe Your Content',
    'Enter a prompt for the new presentation (e.g., "I am so and so and my project entails such and such"):',
    ui.ButtonSet.OK_CANCEL
  );

  // Exit if the user cancels or enters no text
  if (promptResponse.getSelectedButton() !== ui.Button.OK || !promptResponse.getResponseText()) {
    return;
  }
  const promptText = promptResponse.getResponseText();

  // check gemini api key
  if (!GEMINI_API_KEY || GEMINI_API_KEY === 'YOUR_GEMINI_API_KEY_HERE') {
    ui.alert('Configuration Needed', 'Add a gemini api key to the file.', ui.ButtonSet.OK);
    return;
  }

  try {
    const activePresentation = SlidesApp.getActivePresentation();
    const templateId = activePresentation.getId();

    // Provide immediate feedback that the process has started
    ui.alert('Processing your request... This may take a moment. You can close this dialog.');

    // 3. create new copy of this presentation
    Logger.log('Copying template presentation...');
    const newPresentationName = `[Generated] ${activePresentation.getName()} - ${new Date().toLocaleDateString()}`;
    const newFile = DriveApp.getFileById(templateId).makeCopy(newPresentationName);
    const newPresentationId = newFile.getId();
    Logger.log(`New presentation created with ID: ${newPresentationId}`);

    // 4. get JSON of template presentation
    Logger.log('Fetching presentation data...');
    const templateJson = Slides.Presentations.get(templateId);
    //const exampleJson = Slides.Presentations.get(examplePresentationId); --Commented out because of testing with only prompt
    Logger.log('Successfully fetched presentation JSON.');

    // 5. Parse template for text to shorten token input
    let textToReplace = [];
    const presentationJson = templateJson;
    if (presentationJson && presentationJson.slides) {
      presentationJson.slides.forEach(slide => {
        textToReplace.push('Slide number: ' + (slide.objectId || 'unknown'));
        if (slide.pageElements) {
          slide.pageElements.forEach(element => {
            if (element.shape && element.shape.text && element.shape.text.textElements) {
              element.shape.text.textElements.forEach(textElement => {
                if (textElement.textRun && textElement.textRun.content) {
                  // Add the content to our list of text to potentially replace
                  // You might want to add context like slide ID or element ID if needed
                  textToReplace.push(textElement.textRun.content.trim());
                }
              });
            }
          });
        }
      });
    }
    console.log(textToReplace);

    // 6. Send prompt to gemini
    const geminiPrompt = `
      Given a json representation of a presentation from google slides api, create a batch update json body request to update the presentation given the following information:
      USER PROMPT (The specific information to use):
      "${promptText}"

      The batchupdate json contains a replaceAllText field. Given this template and its text, replace ALL text found in the "content" fields with relevant information through batchupdate. 
            Each slide is an object in this representation, 
            so the content is relevant to each other. Replace all texts for all slides. Do it as best as you can. Generate ONLY the JSON array for the 'requests' field of the presentations.batchUpdate API call.
            Do NOT include any surrounding text, explanations, or markdown code blocks. This should be a json object that is an array of "requests": Here is the an array of all the text in the slides:
            ${textToReplace}
    `;

    // 7. call the Gemini API.
    Logger.log('Calling Gemini API...');
    const geminiApiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${GEMINI_API_KEY}`;
    const geminiResponse = UrlFetchApp.fetch(geminiApiUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ contents: [{ parts: [{ text: geminiPrompt }] }] }),
      muteHttpExceptions: true,
    });

    const geminiResult = JSON.parse(geminiResponse.getContentText());
    if (geminiResponse.getResponseCode() !== 200 || !geminiResult.candidates || !geminiResult.candidates[0].content.parts[0].text) {
        throw new Error(`Failed to get a valid response from the Gemini API. Response: ${JSON.stringify(geminiResult)}`);
    }
    Logger.log('Successfully received response from Gemini API.');

    // 8. Parse batchupdate request
    const batchUpdateJsonString = geminiResult.candidates[0].content.parts[0].text;
    const cleanedJsonString = batchUpdateJsonString.replace(/```json/g, '').replace(/```/g, '').trim();
    const batchUpdateRequest = { requests: JSON.parse(cleanedJsonString) };

    // 9. send request to slides api
    Logger.log('Applying updates to the new presentation...');
    Slides.Presentations.batchUpdate(batchUpdateRequest, newPresentationId);
    Logger.log('Updates applied successfully.');

    // 10. display new link
    const newPresentationUrl = `https://docs.google.com/presentation/d/${newPresentationId}/`;
    const message = HtmlService.createHtmlOutput(`<b>Success!</b><br><br>Your new presentation is ready. <a href="${newPresentationUrl}" target="_blank">Click here to open it.</a>`)
        .setWidth(350)
        .setHeight(100);
    ui.showModalDialog(message, 'Generation Complete');

  } catch (err) {
    Logger.log(err.stack);
    ui.alert('An Error Occurred', `An error occurred during the process. Please check the logs for more details. Error: ${err.message}`, ui.ButtonSet.OK);
  }
}
