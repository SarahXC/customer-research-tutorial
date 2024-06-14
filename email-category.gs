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
