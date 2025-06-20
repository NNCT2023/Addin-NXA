### Trước tiên bạn cần lấy [API keys in Google Cloud](https://makersuite.google.com/app/apikey)


```
const API_KEY = 'YOUR_API_KEY'; 
const URL = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=${API_KEY}`;

function askGPT(prompt) {
  const payload = {
    "contents":[
        {"role": "model",
         "parts":[{
           "text": "You are a helpful assisstant.  Yor name is J2TEAM GPT"}]},
        {"role": "user",
         "parts":[{
           "text": prompt }]
        },
    ]
  }

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  };

  try {
    const response = UrlFetchApp.fetch(URL, options);
    const json = JSON.parse(response.getContentText());
    return json["candidates"][0]["content"]["parts"][0]["text"];
  } catch (e) {
    return `Error: ${e.message}`;
  }
}
```
// Author: [JUNO_OKYO - J2TEAM](https://gist.github.com/J2TEAM/91bc15ab65941d8db9cfb30de2b849a3)  
// Recreated by [orinn2k7](https://gist.github.com/orius2k7/21bb709118e264e48277f66fd222a70d)
