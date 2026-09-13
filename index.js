require("dotenv").config();
const express = require("express");
const {
  CloudAdapter,
  ConfigurationBotFrameworkAuthentication,
  TeamsActivityHandler,
  MessageFactory
} = require("botbuilder");

// Bot Framework認証設定
const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication({
  MicrosoftAppId: process.env.MicrosoftAppId,
  MicrosoftAppPassword: process.env.MicrosoftAppPassword,
  MicrosoftAppType: "MultiTenant"
});

const adapter = new CloudAdapter(botFrameworkAuthentication);

adapter.onTurnError = async (context, error) => {
  console.error(`[onTurnError] ${error}`);
  await context.sendActivity("処理中にエラーが発生しました。");
};

// Bot FrameworkのトークンでTeamsの画像を取得
async function getBotFrameworkToken() {
  const res = await fetch(
    "https://login.microsoftonline.com/botframework.com/oauth2/v2.0/token",
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        grant_type: "client_credentials",
        client_id: process.env.MicrosoftAppId,
        client_secret: process.env.MicrosoftAppPassword,
        scope: "https://api.botframework.com/.default"
      })
    }
  );
  const data = await res.json();
  return data.access_token;
}

async function downloadTeamsImage(contentUrl) {
  const token = await getBotFrameworkToken();
  const res = await fetch(contentUrl, { headers: { Authorization: `Bearer ${token}` } });
  const buffer = await res.arrayBuffer();
  return Buffer.from(buffer).toString("base64");
}

// Azure OpenAI Serviceで画像解析
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

// ボットのロジック本体
class PhotoAnalysisBot extends TeamsActivityHandler {
  constructor() {
    super();
    this.onMessage(async (context, next) => {
      const attachments = context.activity.attachments || [];
      const imageAttachment = attachments.find(a => a.contentType?.startsWith("image/"));
      const commentText = context.activity.text?.trim();

      if (imageAttachment) {
        const imageBase64 = await downloadTeamsImage(imageAttachment.contentUrl);
        const replyText = await analyzeImage(
          imageBase64,
          imageAttachment.contentType,
          commentText || "この画像の内容について説明してください"
        );
        await context.sendActivity(MessageFactory.text(replyText));
      }

      await next();
    });
  }
}

const bot = new PhotoAnalysisBot();

const app = express();
app.use(express.json());

app.post("/api/messages", async (req, res) => {
  await adapter.process(req, res, (context) => bot.run(context));
});

const port = process.env.PORT || 3978;
app.listen(port, () => console.log(`Bot listening on port
