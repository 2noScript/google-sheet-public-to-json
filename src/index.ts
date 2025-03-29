// Google Sheet URLs
// https://docs.google.com/spreadsheets/d/1SnJOpCw7LVVfarTkGKJQRSCWR3xJgPEqHKILjRgF_oc/edit?gid=0#gid=0
// https://docs.google.com/spreadsheets/d/1SnJOpCw7LVVfarTkGKJQRSCWR3xJgPEqHKILjRgF_oc/edit?usp=sharing

/**
 * Extracts the sheet ID from a Google Sheets URL
 * @param url Google Sheets URL
 * @returns The sheet ID
 */
function extractSheetId(url: string): string {
  const regex = /\/d\/([a-zA-Z0-9-_]+)/;
  const match = url.match(regex);
  
  if (!match || match.length < 2) {
    throw new Error('Invalid Google Sheets URL');
  }
  
  return match[1];
}

/**
 * Fetches Google Sheets data and converts it to JSON
 * @param url Google Sheets URL
 * @returns Promise with the JSON data
 */
async function convertSheetToJson(url: string): Promise<any> {
  try {
    const sheetId = extractSheetId(url);
    const apiUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:json`;

    const response = await fetch(apiUrl);
    
    if (!response.ok) {
      throw new Error(`Failed to fetch sheet data: ${response.statusText}`);
    }
    
    const text = await response.text();
    const jsonData = JSON.parse(text.substring(text.indexOf('(') + 1, text.lastIndexOf(')')));
    return jsonData;
  } catch (error) {
    console.error('Error converting sheet to JSON:', error);
    throw error;
  }
}

/**
 * Processes multiple Google Sheet URLs and returns their JSON data
 * @param urls Array of Google Sheet URLs
 * @returns Promise with an array of JSON data
 */
async function processSheetUrls(urls: string[]): Promise<any[]> {
  try {
    const promises = urls.map(url => convertSheetToJson(url));
    return await Promise.all(promises);
  } catch (error) {
    console.error('Error processing sheet URLs:', error);
    throw error;
  }
}

// Example usage
const sheetUrls = [
  'https://docs.google.com/spreadsheets/d/1-DBXaIoapJZOrVxcBmYEntfGQXBpqquBcbErPoeTCA4/edit?gid=0#gid=0'
];

// Process the URLs
processSheetUrls(sheetUrls)
  .then(results => {
    console.log('Sheet data as JSON:', results);
  })
  .catch(error => {
    console.error('Failed to process sheets:', error);
  });