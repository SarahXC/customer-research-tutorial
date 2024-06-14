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
