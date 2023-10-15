```javascript
import sheetPublicToJson from "google-sheet-public-to-json";

const data = await sheetPublicToJson(
  "https://docs.google.com/spreadsheets/d/{sheetId}/edit#gid=0"
);
console.log(data);
```
