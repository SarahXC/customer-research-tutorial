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
