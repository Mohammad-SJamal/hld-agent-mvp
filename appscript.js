//this file is used to create a Google Slides add-on that generates a presentation based on a user prompt and a template presentation.
// It uses the Gemini API to generate content based on the provided prompt and updates the template presentation with the generated content.


const scriptProperties = PropertiesService.getScriptProperties();
const GEMINI_API_KEY = scriptProperties.getProperty("ai_key"); 
//this will fetch the API key from the script properties.
//For more info --> https://developers.google.com/apps-script/guides/properties

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
  
  // 1. prompt the user for the core information to generate the presentation.
  const promptResponse = ui.prompt(
    'Step 2: Describe Your Content',
    'Enter a prompt for the new presentation (e.g., "I am so and so and my project entails such and such"):',
    ui.ButtonSet.OK_CANCEL
  );

  // exit if the user cancels or enters no text
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
                  textToReplace.push(textElement.textRun.content.trim());
                }
              });
            }
          });
        }
      });
    }
    console.log(textToReplace);


    // 6. Create prompt for gemini
    const old_geminiPrompt = `
      Given a json representation of a presentation from google slides api, create a batch update 
      json body request to update the presentation given the following information:
      USER PROMPT (The specific information to use):
      "${promptText}"

      The batchupdate json contains a replaceAllText field. Given this template and its text, 
      replace ALL text found in the "content" fields with relevant information through batchupdate. 
            Each slide is an object in this representation, 
            so the content is relevant to each other. Replace all texts for all slides. 
            Do it as best as you can. Generate ONLY the JSON array for the 'requests' 
            field of the presentations.batchUpdate API call.
            Do NOT include any surrounding text, explanations, or markdown code blocks. 
            This should be a json object that is an array of "requests".
            Replace the date in the presentation as ${Utilities.formatDate(new Date(), "GMT+1", "MM/yyyy")}.  
            Here is the an array of all the text in the slides:
            ${textToReplace}
    `;

    const geminiPrompt = `
      You are a robot assisting a solution architect in getting a jump start on a new high level design google slides presentation from their 
      details and a template google slide presentation, in addition to high level design slide documents in the same product area.
      "${promptText}"
      Your task is to read the text of the json and generate a batchUpdate json from google slides API to replace each string shown in the array with
      relevant information. 
      The batchupdate json contains a replaceAllText field. Given this template's text, 
      replace ALL text with relevant information through batchupdate. Additionally, add a new slide regarding the specific technology/technique to be implemented. 
      Generate ONLY the JSON array for the 'requests' field of the presentations.batchUpdate API call.
      The JSON must be syntactically perfect and strictly adhere to the Google Slides API's request formats.

      For existing slides, use replaceAllText requests to update the content.

     
      Do NOT include any surrounding text, explanations, or markdown code blocks.
      Replace the date in the presentation as ${Utilities.formatDate(new Date(), "GMT+1", "MM/yyyy")}. 
      This should be a valid json that is an array of "requests": Here is the an array of all the text in the slides:
      ${textToReplace}
    `;

//  For creating a new slide:
//       1.  Use a createSlide request with a unique objectId (e.g., "newSlideXYZ") and specify slideLayoutReference as "TITLE_AND_BODY".
//       2.  Immediately following the createSlide request, add two insertText requests.
//           * The first insertText request should populate the *title* of the new slide. Assign a unique objectId for this text box (e.g., "newSlideXYZ_title").
//           * The second insertText request should populate the *body* of the new slide. Assign a unique objectId for this text box (e.g., "newSlideXYZ_body").
//           Ensure these insertText requests correctly reference the objectId assigned to the *text box* within the newly created slide, not the slide's own objectId for the insertText's objectId field.


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

    // 8.5 Debug to batchUpdate generate file
    DriveApp.createFile(`New Json ${Utilities.formatDate(new Date(), "America/Chicago", "HH:mm:ss")}`, batchUpdateRequest);

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
