// Copyright (C) 2020 Potix Corporation. All Rights Reserved
// History: 2020/12/22 5:55 PM
// Author: charlie<charliehsieh@potix.com>

const path = require('path');
const express = require('express');
const ENV_FILE = path.join(__dirname, '../boeneo/mis/configurations/config/msteams/env');
require('dotenv').config({ path: ENV_FILE });
const { BotFrameworkAdapter } = require('botbuilder');
const { MicrosoftAppCredentials } = require('botframework-connector');
const { BotActivityHandler } = require('./bot/botActivityHandler');
const { QuireApi } = require('./utils/quireApi');
const quireNotificationHandler = require('./bot/quireNotificationHandler');
const dbAccess = require('./db/dbAccess');

// Create adapter.
const adapter = new BotFrameworkAdapter({
  appId: process.env.BotId,
  appPassword: process.env.BotPassword
});

adapter.onTurnError = async (context, error) => {
  console.error(`\n [onTurnError] unhandled error: ${error}`);

  if (error.isAxiosError) {
    console.log(error.config);
    const statusCode = error.response.status;
    if (statusCode === 429 || statusCode === 503) {
      await context.sendActivity('Service is unavailable, please try again later');
      return;
    }
  }

  await context.sendActivity('Internal error');
};

// Create bot handlers
const botActivityHandler = new BotActivityHandler();

// Create HTTP server.
const app = express();
const port = process.env.port || process.env.PORT || 3978;
app.use(express.json());
const server = app.listen(port, () =>
  console.log(`\Bot/ME service listening at https://localhost:${port}`)
);

// Listen for incoming requests.
app.post('/api/messages', (req, res) => {
  adapter.processActivity(req, res, async (context) => {
    // Process bot activity
    await botActivityHandler.run(context);
  });
});

// Handle notifications
app.post('/webhook*', async (req, res) => {
  const conversationId = req.path.substring(9);
  const serviceUrl = req.body.channel;
  const ref = {
    conversation: {
      id: conversationId
    },
    serviceUrl: serviceUrl
  }

  MicrosoftAppCredentials.trustServiceUrl(serviceUrl);
  adapter.continueConversation(ref, async context => {
    try {
      await quireNotificationHandler.handleQuireNotification(context, req.body.data);
      res.sendStatus(200);
    } catch (error) {
      if (error.statusCode === 403) {
        console.log(error.body);
        res.sendStatus(403);
      } else {
        console.log(error.config);
        res.sendStatus(200);
      }
    }
  });
});

// Handle login
app.get('/bot-auth-start', (req, res) => QuireApi.handleAuthStart(req, res));
app.get('/bot-auth-end', (req, res) => QuireApi.handleAuthEnd(req, res));

// heartbeat
app.get('/heartbeat', (req, res) => res.sendStatus(200));

// graceful shutdown
process.on('SIGINT', () => {
  dbAccess.shutdown();
  server.close();
  console.log(`\nServer shutdown...`);
});
