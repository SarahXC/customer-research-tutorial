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
