async function updateCompanyDescription() {
  console.log('Updating company description...');
  const rows = worksheet.getDataRange().getValues();
  
  for (let i = 1; i < rows.length; i++) {
    const companyWebsite = rows[i][companyWebsiteColumnIndex];
    
    if (companyWebsite) {
      try {
        const companyDescription = await fetchCompanyDescriptionFromExa(companyWebsite);
        const summarizedCompanyDescription = await summarizeCompanyDescription(companyDescription);
        worksheet.getRange(i + 1, companyDescriptionColumnIndex + 1).setValue(summarizedCompanyDescription);
      } catch (error) {
        console.error(`Error updating company description for row ${i}:`, error);
      }
    }
  }
}

/**
 * Fetch the company description from the EXA API.
 * @param {string} websiteUrl - The URL of the company's website.
 * @returns {string} - The description of the company.
 */
async function fetchCompanyDescriptionFromExa(websiteUrl) {
  const apiUrl = 'https://api.exa.ai/contents';
  const requestPayload = {
    ids: [websiteUrl]
  };

  const requestOptions = {
    method: 'POST',
    contentType: 'application/json',
    headers: {
      'x-api-key': EXA_API_KEY
    },
    payload: JSON.stringify(requestPayload),
    muteHttpExceptions: true
  };

  try {
    const response = await UrlFetchApp.fetch(apiUrl, requestOptions);
    const data = JSON.parse(response.getContentText());

    if (data.results && data.results.length > 0) {
      return data.results[0].text;
    }
  } catch (error) {
    console.error('Error fetching company description from EXA API:', error);
  }

  return '';
}

/**
 * Summarize the company description using OpenAI's GPT-3.5 Turbo model.
 * @param {string} companyDescription - The full text of the company description.
 * @returns {string} - A summarized description of the company.
 */
async function summarizeCompanyDescription(companyDescription) {
  const systemMessage = "You are a helpful assistant given the full text contents of a company's homepage. Your task is to write a short, 2-3 sentence summary of what the company does. Start the summary with the name of the company. For example, Exa AI is a...";
  
  // If the company description is empty, return an empty string
  if (!companyDescription) {
    return "";
  }

  const apiKey = OPENAI_API_KEY;
  const apiUrl = "https://api.openai.com/v1/chat/completions";

  const requestBody = {
    model: "gpt-3.5-turbo",
    messages: [
      { role: "system", content: systemMessage },
      { role: "user", content: companyDescription }
    ],
    temperature: 0
  };

  const requestOptions = {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + apiKey
    },
    payload: JSON.stringify(requestBody),
    muteHttpExceptions: true
  };

  try {
    const response = await UrlFetchApp.fetch(apiUrl, requestOptions);
    const responseJson = JSON.parse(response.getContentText());
    return responseJson.choices[0].message.content;
  } catch (error) {
    console.error('Error summarizing company description:', error);
  }

  return '';
}
