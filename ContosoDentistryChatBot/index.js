// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const path = require('path');

const dotenv = require('dotenv');
// Import required bot configuration.
const ENV_FILE = path.join(__dirname, '.env');
dotenv.config({ path: ENV_FILE });

const restify = require('restify');

// Import required bot services.
// See https://aka.ms/bot-services to learn more about the different parts of a bot.
const { BotFrameworkAdapter } = require('botbuilder');

// This bot's main dialog.
const { DentaBot } = require('./bot');

// Create HTTP server.
const server = restify.createServer();
const port = process.env.PORT || process.env.port || 3978;

server.listen(port, () => {
    const address = server.address();
    const host = address.address === '::' ? 'localhost' : address.address;
    console.log(`\n${server.name} listening to http://${host}:${address.port}`);
    console.log('\nGet Bot Framework Emulator: https://aka.ms/botframework-emulator');
    console.log('\nTo talk to your bot, open the emulator select "Open Bot"');
});

// Create adapter.
// See https://aka.ms/about-bot-adapter to learn more about how bots work.
const adapter = new BotFrameworkAdapter({
    appId: process.env.MicrosoftAppId,
    appPassword: process.env.MicrosoftAppPassword
});

// Catch-all for errors.
const onTurnErrorHandler = async (context, error) => {
    // This check writes out errors to console log .vs. app insights.
    // NOTE: In production environment, you should consider logging this to Azure
    //       application insights.
    console.error(`\n [onTurnError] unhandled error: ${error}`);

    // Send a trace activity, which will be displayed in Bot Framework Emulator
    await context.sendTraceActivity(
        'OnTurnError Trace',
        `${error}`,
        'https://www.botframework.com/schemas/error',
        'TurnError'
    );

    // Send a message to the user
    await context.sendActivity('The bot encountered an error or bug.');
    await context.sendActivity('To continue to run this bot, please fix the bot source code.');
};

// Set the onTurnError for the singleton BotFrameworkAdapter.
adapter.onTurnError = onTurnErrorHandler;

// Map configuration values values from .env file into the required format for each service.
const QnAConfiguration = {
    projectName: process.env.QnAProjectName || process.env.QnAKnowledgebaseId,
    deploymentName: process.env.QnADeploymentName || 'production',
    endpointKey: process.env.QnAEndpointKey || process.env.QnAAuthKey,
    host: process.env.QnAEndpointHostName
};

const CLUConfiguration = {
    projectName: process.env.CLUProjectName,
    deploymentName: process.env.CLUDeploymentName,
    endpointKey: process.env.CLUEndpointKey,
    host: process.env.CLUEndpointHostName
};

const LuisConfiguration = {
    applicationId: process.env.LuisAppId,
    endpointKey: process.env.LuisAPIKey,
    endpoint: process.env.LuisAPIHostName
};

const SchedulerConfiguration = {
    SchedulerEndpoint: process.env.SchedulerEndpoint || process.env.SchedulerApiUrl
};

// Pack each service configuration into one object.
const configuration = {
    QnAConfiguration,
    CLUConfiguration,
    LuisConfiguration,
    SchedulerConfiguration
};

// Create the main dialog.
const myBot = new DentaBot(configuration, {});

// Listen for incoming requests.
server.post('/api/messages', async (req, res) => {
    try {
        await adapter.processActivity(req, res, async (context) => {
            // Route to main dialog.
            await myBot.run(context);
        });
    } catch (error) {
        console.error(`\n [server.post /api/messages] error: ${error}`);
        if (!res.headersSent) {
            res.statusCode = 500;
            res.end();
        }
    }
});

// Listen for Upgrade requests for Streaming.
server.on('upgrade', (req, socket, head) => {
    // Create an adapter scoped to this WebSocket connection to allow storing session data.
    const streamingAdapter = new BotFrameworkAdapter({
        appId: process.env.MicrosoftAppId,
        appPassword: process.env.MicrosoftAppPassword
    });
    // Set onTurnError for the BotFrameworkAdapter created for each connection.
    streamingAdapter.onTurnError = onTurnErrorHandler;

    streamingAdapter.useWebSocket(req, socket, head, async (context) => {
        // After connecting via WebSocket, run this logic for every request sent over
        // the WebSocket connection.
        await myBot.run(context);
    });
});
