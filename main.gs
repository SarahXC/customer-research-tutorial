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
