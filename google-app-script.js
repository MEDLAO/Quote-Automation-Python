function generateQuoteFromSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('GroupedQuotes');
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const rows = data.slice(1);  // Skip header

  rows.forEach(row => {
    const rowData = Object.fromEntries(header.map((key, i) => [key, row[i]]));
    const servicesJson = rowData['Services'];

    try {
      const services = JSON.parse(servicesJson);  // Parse the JSON string

      // Create a new Google Doc
      const doc = DocumentApp.create(`Quote - ${rowData['Quote ID']}`);
      const body = doc.getBody();

      body.appendParagraph("PROFESSIONAL SERVICE QUOTATION");
      body.appendParagraph(`Quote ID: ${rowData['Quote ID']}`);
      body.appendParagraph(`Date: ${rowData['Date']}`);
      body.appendParagraph(`Client: ${rowData['Client Name']}`);
      body.appendParagraph(`Email: ${rowData['Email']}`);
      body.appendParagraph(`Organization: ${rowData['Organization']}`);
      body.appendParagraph("\nSERVICES INCLUDED:\n");

      // Inject each service
      services.forEach(service => {
        body.appendParagraph(`Service Type: ${service.Service_Type}`);
        body.appendParagraph(`Language Pair: ${service.Language_Pair}`);
        body.appendParagraph(`Modality: ${service.Modality}`);
        body.appendParagraph(`Word Count: ${service.Word_Count}`);
        body.appendParagraph(`Duration (hrs): ${service.Duration_hrs}`);
        body.appendParagraph(`Rate: ${service.Rate}`);
        body.appendParagraph(`Details: ${service.Details}`);
        body.appendParagraph(`Total: ${service.Total} USD`);
        body.appendParagraph('---');
      });

      body.appendParagraph(`\nGrand Total: ${rowData['Grand Total']} USD`);
      body.appendParagraph(`\nNotes:\n${rowData['Notes'] || ''}`);

      Logger.log(`Created document: ${doc.getUrl()}`);

    } catch (err) {
      Logger.log(`Error parsing Services JSON for Quote ID ${rowData['Quote ID']}: ${err}`);
    }
  });
}
