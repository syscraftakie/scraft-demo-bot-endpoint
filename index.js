async function analyzeImage(imageBase64, mediaType, userComment) {
  const endpoint = process.env.AZURE_OPENAI_ENDPOINT;
  const deployment = process.env.AZURE_OPENAI_DEPLOYMENT;
  const apiKey = process.env.AZURE_OPENAI_API_KEY;

const response = await fetch(
    `${endpoint}/openai/deployments/${deployment}/chat/completions?api-version=2025-01-01-preview`,
    {
      method: "POST",
      headers: {
        "api-key": apiKey,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        messages: [
          {
            role: "user",
            content: [
              { type: "text", text: userComment },
              {
                type: "image_url",
                image_url: { url: `data:${mediaType};base64,${imageBase64}` }
              }
            ]
          }
        ],
        max_tokens: 1024
      })
    }
  );

  const data = await response.json();
  return data.choices[0].message.content;
}