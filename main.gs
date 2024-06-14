async function enrichSheet() {
  await updateEmailCategory();
  await updateCompanyWebsite();
  await updatePersonalLinkedin();
  await updateCompanyLinkedin();
  await updateCompanyDescription();
  await updateCompanyCategory();
  await updateRouting();
  await updateEmailDraft();
}

// SETUP and load the spreadsheet
const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
const worksheet = spreadsheet.getSheetByName('Sheet1');
const headers = worksheet.getDataRange().getValues()[0];

// Get the heading indicies
const nameColumnIndex = headers.findIndex(header => header.toLowerCase() === 'name');
const emailColumnIndex = headers.findIndex(header => header.toLowerCase() === 'email');
const emailCategoryColumnIndex = headers.findIndex(header => header.toLowerCase() === 'email category');
const companyWebsiteColumnIndex = headers.findIndex(header => header.toLowerCase() === 'company website');
const companyDescriptionColumnIndex = headers.findIndex(header => header.toLowerCase() === 'company description');
const personalLinkedinColumnIndex = headers.findIndex(header => header.toLowerCase() === 'personal linkedin');
const companyLinkedinColumnIndex = headers.findIndex(header => header.toLowerCase() === 'company linkedin');
const companyCategoryColumnIndex = headers.findIndex(header => header.toLowerCase() === 'company category');
const routingColumnIndex = headers.findIndex(header => header.toLowerCase() === 'routing');
const emailDraftColumnIndex = headers.findIndex(header => header.toLowerCase() === 'email draft');

// Read the templates
const [categoryDescriptions, emailTemplates, routingTemplates] = readTemplates();
const categoriesString = Object.entries(categoryDescriptions)
  .map(([category, description], index) => `Category ${index + 1}: ${category}\nDescription: ${description}`)
  .join('\n');

function readTemplates() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const worksheet = spreadsheet.getSheetByName('templates');
  const rows = worksheet.getDataRange().getValues();

  const categoryDescriptions = {};
  const emailTemplates = {};
  const routingTemplates = {};

  for (let i = 1; i < rows.length; i++) {
    const [category, description, emailTemplate, routingTemplate] = rows[i];
    categoryDescriptions[category] = description;
    emailTemplates[category] = emailTemplate;
    routingTemplates[category] = routingTemplate;
  }
  return [categoryDescriptions, emailTemplates, routingTemplates];
}

// Email category
async function updateEmailCategory() {
  console.log('Updating email category...');
  try {
    const rows = worksheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const email = rows[i][emailColumnIndex];

      const category = determineEmailCategory(email);
      worksheet.getRange(i + 1, emailCategoryColumnIndex + 1).setValue(category);
    }
    console.log('Email categories updated successfully');
    return 'Email categories updated successfully';
  } catch (error) {
    console.error(`Error updating email categories: ${error.message}`);
    throw new Error(`Error updating email categories: ${error.message}`);
  }
}

/**
 * Determines the category of an email based on its domain.
 * @param {string} email - The email address to categorize.
 * @returns {string} - The category of the email.
 */
function determineEmailCategory(email) {
  const personalDomains = [
    'yahoo.com', 'qq.me', 'qq.com', 'duck.com', 'gmail.com', 'google.com', 
    'hotmail.com', 'proton.me', 'googlemail.com', 'icloud.com', 'outlook.com'
  ];
  
  if (email.includes('edu')) {
    return 'Student';
  } else if (personalDomains.some(domain => email.endsWith(domain))) {
    return 'Personal';
  } else {
    return 'Company';
  }
}

// Company website
async function updateCompanyWebsite() {
  console.log('Updating company website...');
  const rows = worksheet.getDataRange().getValues();
  
  for (let i = 1; i < rows.length; i++) {
    const email = rows[i][emailColumnIndex];
    const emailCategory = rows[i][emailCategoryColumnIndex];

    if (emailCategory === 'Company') {
      try {
        const companyWebsite = await fetchCompanyWebsiteFromEmail(email);
        console.log({ rowIndex: i, companyWebsite });
        worksheet.getRange(i + 1, companyWebsiteColumnIndex + 1).setValue(companyWebsite);
      } catch (error) {
        console.error(`Error updating company website for row ${i}:`, error);
      }
    }
  }
}

async function fetchCompanyWebsiteFromEmail(email) {
  const emailDomain = email.split('@')[1];
  const searchQuery = `${emailDomain} company`;

  const apiUrl = 'https://api.exa.ai/search';
  const requestPayload = {
    query: searchQuery,
    use_autoprompt: false,
    num_results: 1,
    type: 'keyword',
    exclude_domains: [
      'linkedin.com', 
      'twitter.com', 
      'wikipedia.org', 
      'ycombinator.com', 
      'github.com'
    ]
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
      const firstResult = data.results[0];
      return firstResult.url;
    }
  } catch (error) {
    console.error('Error retrieving company URL from EXA API:', error);
  }

  return 'unknown';
}

// Personal Linkedin 
async function updatePersonalLinkedin() {
  console.log('Updating personal Linkedin...');
  const rows = worksheet.getDataRange().getValues();
  
  for (let i = 1; i < rows.length; i++) {
    const name = rows[i][nameColumnIndex];
    const email = rows[i][emailColumnIndex];
    const emailCategory = rows[i][emailCategoryColumnIndex];

    if (emailCategory === 'Company') {
      try {
        const personalLinkedin = await fetchPersonalLinkedinFromNameAndEmail(name, email);
        worksheet.getRange(i + 1, personalLinkedinColumnIndex + 1).setValue(personalLinkedin);
      } catch (error) {
        console.error(`Error updating personal LinkedIn for row ${i}:`, error);
      }
    }
  }
}

async function fetchPersonalLinkedinFromNameAndEmail(name, email) {
  const emailDomain = email.split('@')[1];
  const searchQuery = `${name} ${emailDomain}`;

  const apiUrl = 'https://api.exa.ai/search';
  const requestPayload = {
    query: searchQuery,
    use_autoprompt: false,
    num_results: 1,
    type: 'keyword',
    include_domains: ['linkedin.com']
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
      const firstResult = data.results[0];
      if (firstResult.url.includes('linkedin.com/posts')) {
        return 'unknown';
      }
      return firstResult.url;
    }
  } catch (error) {
    console.error('Error retrieving Linkedin URL from EXA API:', error);
  }

  return 'unknown';
}


// Company Linkedin
async function updateCompanyLinkedin() {
  console.log('Updating company Linkedin...');
  const rows = worksheet.getDataRange().getValues();
  
  for (let i = 1; i < rows.length; i++) {
    const email = rows[i][emailColumnIndex];
    const emailCategory = rows[i][emailCategoryColumnIndex];

    if (emailCategory === 'Company') {
      try {
        const companyLinkedin = await fetchCompanyLinkedinFromDomain(email);
        console.log(companyLinkedin)
        worksheet.getRange(i + 1, companyLinkedinColumnIndex + 1).setValue(companyLinkedin);
      } catch (error) {
        console.error(`Error updating company Linkedin for row ${i}:`, error);
      }
    }
  }
}

async function fetchCompanyLinkedinFromDomain(email) {
  const emailDomain = email.split('@')[1];

  const apiUrl = 'https://api.exa.ai/search';
  const requestPayload = {
    query: emailDomain,
    use_autoprompt: false,
    num_results: 5,
    type: 'keyword',
    include_domains: ['linkedin.com']
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
      for (const result of data.results) {
        if (result.url.includes('linkedin.com/company')) {
          return result.url;
        }
      }
      return 'unknown';
    }
  } catch (error) {
    console.error('Error retrieving company Linkedin URL from EXA API:', error);
  }

  return 'unknown';
}

// Company description
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

// Company Category 
async function updateCompanyCategory() {
  console.log('Updating company category...');
  try {
    const rows = worksheet.getDataRange().getValues();
    
    for (let i = 1; i < rows.length; i++) {
      const email = rows[i][emailColumnIndex];
      const companyDescription = rows[i][companyDescriptionColumnIndex];
      const companyWebsite = rows[i][companyWebsiteColumnIndex];

      const initialCategory = getCategoryFromEmail(email);
      
      if (initialCategory === 'Company') {
        const companyCategory = await assignCategory(companyDescription, companyWebsite);
        worksheet.getRange(i + 1, companyCategoryColumnIndex + 1).setValue(companyCategory.category);
      } else {
        worksheet.getRange(i + 1, companyCategoryColumnIndex + 1).setValue(initialCategory);
      }
    }
    return 'Company categories updated successfully';
  } catch (error) {
    console.error(`Error updating company categories: ${error.message}`);
    throw new Error(`Error updating company categories: ${error.message}`);
  }
}

/**
 * Determines the category of an email based on its domain.
 * @param {string} email - The email address to categorize.
 * @returns {string} - The category of the email.
 */
function getCategoryFromEmail(email) {
  const personalDomains = [
    'yahoo.com', 'qq.me', 'qq.com', 'duck.com', 'gmail.com', 'google.com', 
    'hotmail.com', 'proton.me', 'googlemail.com', 'icloud.com', 'outlook.com'
  ];
  
  if (email.includes('edu')) {
    return 'Student';
  } else if (personalDomains.some(domain => email.endsWith(domain))) {
    return 'Personal';
  } else {
    return 'Company';
  }
}

/**
 * Assigns a category to a company based on its description and website.
 * @param {string} companyDescription - The description of the company.
 * @param {string} companyWebsite - The website of the company.
 * @returns {Object} - The category and reasoning for the category.
 */
async function assignCategory(companyDescription, companyWebsite) {
  const content = `
    You are a sales email agent for an AI startup.

    You will be given information about a company. It will include the company description, and company website. Your task will be to identify which of a large number of categories of customers that a company belongs to, provide a reason for why that is the right category.

    ${categoriesString}

    Here is the information about the person and company:

    Company website link: ${companyWebsite}
    Company description: ${companyDescription}

    Think critically and carefully, please construct the proper JSON response for category and reason.
  `;

  const customFunctions = [
    {
      name: 'determine_customer_category',
      description: 'Determines the company category',
      parameters: {
        type: 'object',
        properties: {
          category: {
            type: 'string',
            enum: [
              'AI-Enabled Writing Assistants',
              'General AI Chatbots or Assistants',
              'RAG and Complex Search',
              'Security and Observability AI',
              'AI for Business Solutions',
              'Creative and Marketing AI',
              'AI-Enabled Legal Assistants',
              'Cloud infrastructure',
              'Other'
            ],
            description: 'Category of the company'
          },
          reasoning: {
            type: 'string',
            description: 'How the category was determined'
          }
        }
      }
    }
  ];

  const requestBody = {
    model: "gpt-3.5-turbo",
    messages: [
      { role: "user", content: content }
    ],
    temperature: 0,
    functions: customFunctions,
    function_call: "auto"
  };

  const requestOptions = {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + OPENAI_API_KEY
    },
    payload: JSON.stringify(requestBody)
  };

  try {
    const response = await UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", requestOptions);
    const responseText = response.getContentText();
    const responseJson = JSON.parse(responseText);
    const jsonResponse = JSON.parse(responseJson.choices[0].message.function_call.arguments);

    return jsonResponse;
  } catch (error) {
    console.error('Error assigning company category:', error);
    throw new Error('Error assigning company category');
  }
}

// Routing 
async function updateRouting() {
  console.log('Updating routing...');
  try {
    const rows = worksheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const companyCategory = rows[i][companyCategoryColumnIndex];
      const route = routingTemplates[companyCategory];

      worksheet.getRange(i + 1, routingColumnIndex + 1).setValue(route);
    }
    return 'Routing updated successfully';
  } catch (error) {
    throw new Error(`Error updating routing: ${error.message}`);
  }
}

// Email Draft 
async function updateEmailDraft() {
  console.log('Updating email draft...');
  try {
    const rows = worksheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const name = rows[i][nameColumnIndex];
      const companyWebsite = rows[i][companyWebsiteColumnIndex];
      const companyDescription = rows[i][companyDescriptionColumnIndex]; 
      const companyCategory = rows[i][companyCategoryColumnIndex];
      const emailTemplate = emailTemplates[companyCategory];

      try {
        const draftedEmail = await draftEmailBasedOnTemplate(name, companyWebsite, companyDescription, companyCategory, emailTemplate);
        if (emailDraftColumnIndex !== -1) {
          worksheet.getRange(i + 1, emailDraftColumnIndex + 1).setValue(draftedEmail);
        }
      } catch (error) {
        console.error(`Error drafting email for row ${i + 1}: ${error.message}`);
      }
    }
    console.log('Email drafts updated successfully');
  } catch (error) {
    console.error('Error updating email drafts:', error);
  }
}

/**
 * Drafts an email based on the provided template and company details.
 * @param {string} name - The name of the person.
 * @param {string} companyWebsite - The website of the company.
 * @param {string} companyDescription - The description of the company.
 * @param {string} companyCategory - The category of the company.
 * @param {string} emailTemplate - The template to be used for drafting the email.
 * @returns {string} - The drafted email content.
 */
async function draftEmailBasedOnTemplate(name, companyWebsite, companyDescription, companyCategory, emailTemplate) {
  const content = `
    You are a sales email agent for an AI startup.

    You will be given information about a person and their company. It will include their name, company description, and company website. Further, previously, you were tasked with categorizing the company into a bucket, so we will also give you this company categorization as well as an email template to use for the email. Note that this email template has been very carefully written, so please only be creative if you are really confident, otherwise follow the template.

    You will draft an email to that company+person that adheres to the template, which is standard and includes instructions/fill-ins in square brackets. Do not include an email signature at the end. For example, do not include anything like \`Best,\`, \`Regards,\` etc. If there is no name, just say \`Hi there!\`

    Name: ${name}
    Company website link: ${companyWebsite}
    Company description: ${companyDescription}
    Company category: ${companyCategory}
    Email template: ${emailTemplate}

    Return only the finished email.
  `;

  const requestBody = {
    model: "gpt-3.5-turbo",
    messages: [
      { role: "user", content: content }
    ],
    temperature: 0
  };

  const requestOptions = {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + OPENAI_API_KEY
    },
    payload: JSON.stringify(requestBody)
  };

  try {
    const response = await UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", requestOptions);
    const responseText = response.getContentText();
    const responseJson = JSON.parse(responseText);

    return responseJson.choices[0].message.content;
  } catch (error) {
    console.error('Error drafting email:', error);
    throw new Error('Error drafting email');
  }
}
