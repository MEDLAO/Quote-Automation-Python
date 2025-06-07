function injectServicesColumn() {
  // Get the sheet named "GroupedQuotes"
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GroupedQuotes');

  // Get all values from the sheet (including the header)
  const data = sheet.getDataRange().getValues();

  // Extract the first row as headers (column names)
  const headers = data[0];

  // Find the index of the "Services" column
  const servicesColIndex = headers.indexOf('Services');

  // Loop through each data row (skip the header)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    // Get the raw JSON string from the "Services" cell
    const rawJson = row[servicesColIndex];

    let parsedServices;
    try {
      // Try to parse the JSON string into an array of service objects
      parsedServices = JSON.parse(rawJson);
    } catch (e) {
      // If it's not valid JSON, log a message and skip this row
      Logger.log(`Invalid JSON in row ${i + 1}`);
      continue;
    }

    // Format the service data for display (e.g., one bullet per service)
    const formatted = parsedServices.map(service => {
      return `• ${service.Service_Type} | ${service.Language_Pair} | ${service.Modality} | ${service.Word_Count} words | ${service.Duration_hrs} hrs | ${service.Rate} USD | Total: ${service.Total} USD\n${service.Details}`;
    }).join('\n\n'); // Separate services with two new lines

    // Write the formatted text into a new column to the right of existing data
    // If the sheet has N columns, this writes to column N+1
    sheet.getRange(i + 1, headers.length + 1).setValue(formatted);
  }

  // You can now add a marker like {{Formatted Services}} in your template
  // and Document Studio will use the new column's content.
}
