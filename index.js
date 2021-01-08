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
const quireNotificationHandler = require('./bot/quireNotificationHandler');
const { QuireApi } = require('./utils/quireApi');

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
const server = express();
const port = process.env.port || process.env.PORT || 3978;
server.use(express.json());
server.listen(port, () =>
  console.log(`\Bot/ME service listening at https://localhost:${port}`)
);

// Listen for incoming requests.
server.post('/api/messages', (req, res) => {
  adapter.processActivity(req, res, async (context) => {
    // Process bot activity
    await botActivityHandler.run(context);
  });
});

// Handle notifications
server.post('/webhook*', async (req, res) => {
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
server.get('/bot-auth-start', (req, res) => QuireApi.handleAuthStart(req, res));
server.get('/bot-auth-end', (req, res) => QuireApi.handleAuthEnd(req, res));

// heartbeat
server.get('/heartbeat', (req, res) => res.sendStatus(200));
