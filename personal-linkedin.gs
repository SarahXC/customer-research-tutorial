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
