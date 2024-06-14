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
