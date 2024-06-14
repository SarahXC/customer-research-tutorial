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
