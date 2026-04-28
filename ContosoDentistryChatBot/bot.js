// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ActivityHandler, MessageFactory } = require('botbuilder');

const DentistScheduler = require('./dentistscheduler');

let IntentRecognizer;
try {
    IntentRecognizer = require("./intentrecognizer");
} catch(e) {
    console.error('Error loading IntentRecognizer:', e.message);
    IntentRecognizer = null;
}

class DentaBot extends ActivityHandler {
    constructor(configuration, qnaOptions) {
        // call the parent constructor
        super();
        if (!configuration) throw new Error('[QnaMakerBot]: Missing parameter. configuration is required');

        this.QnAConfiguration = configuration.QnAConfiguration || {};
        this.CLUConfiguration = configuration.CLUConfiguration || {};
       
        // create a DentistScheduler connector
        const schedulerConfig = configuration.SchedulerConfiguration || {};
        const rawSchedulerEndpoint =
            schedulerConfig.SchedulerEndpoint ||
            configuration.SchedulerEndpoint ||
            process.env.SchedulerApiUrl ||
            process.env.SchedulerEndpoint;
        const schedulerEndpoint = rawSchedulerEndpoint && rawSchedulerEndpoint.endsWith('/')
            ? rawSchedulerEndpoint
            : rawSchedulerEndpoint
                ? `${rawSchedulerEndpoint}/`
                : rawSchedulerEndpoint;

        this.scheduler = new DentistScheduler({ SchedulerEndpoint: schedulerEndpoint });
        this.DentistScheduler = this.scheduler;
        console.log(`DentistScheduler initialized with endpoint: ${schedulerEndpoint}`);
      
        // create a IntentRecognizer connector
        this.IntentRecognizer = new IntentRecognizer(configuration.LuisConfiguration);


        this.onMessage(async (context, next) => {
            try {
                // Get user message text
                const userMessage = context.activity.text || '';
                const normalizedMessage = userMessage.toLowerCase();
                console.log('\n--- Incoming Message ---');
                console.log(`User input: "${userMessage}"`);

                let response = null;
                const appointmentTime = normalizedMessage.includes('tomorrow') ? 'tomorrow' : userMessage;

                // Step 1: Detect intent using Azure Conversational Language Understanding
                const cluConfig = this.CLUConfiguration;
                try {
                    if (cluConfig && cluConfig.host && cluConfig.endpointKey && cluConfig.projectName && cluConfig.deploymentName) {
                        const cluHost = cluConfig.host.endsWith('/') ? cluConfig.host : `${cluConfig.host}/`;
                        const cluResponse = await fetch(`${cluHost}language/:analyze-conversations?api-version=2022-05-01`, {
                            method: 'POST',
                            headers: {
                                'Content-Type': 'application/json',
                                'Ocp-Apim-Subscription-Key': cluConfig.endpointKey
                            },
                            body: JSON.stringify({
                                analysisInput: {
                                    conversationItem: {
                                        id: '1',
                                        text: userMessage,
                                        modality: 'text',
                                        language: 'en'
                                    }
                                },
                                parameters: {
                                    projectName: cluConfig.projectName,
                                    deploymentName: cluConfig.deploymentName,
                                    stringIndexType: 'TextElement_V8'
                                },
                                kind: 'Conversation'
                            })
                        });

                        const cluResult = await cluResponse.json();
                        console.log(`CLU API response: ${JSON.stringify(cluResult)}`);
                        const topIntent = cluResult.result && cluResult.result.prediction && cluResult.result.prediction.topIntent;
                        console.log(`Detected intent: ${topIntent}`);

                        if (topIntent === 'GetAvailability') {
                            console.log('Scheduler triggered');
                            console.log('Calling DentistScheduler.getAvailability()...');
                            response = await this.scheduler.getAvailability();
                            console.log(`Scheduler API response: ${response}`);
                            await context.sendActivity(response);
                            await next();
                            return;
                        }

                        if (topIntent === 'ScheduleAppointment') {
                            console.log('Scheduler triggered');
                            console.log('Calling DentistScheduler.scheduleAppointment()...');
                            response = await this.scheduler.scheduleAppointment(appointmentTime);
                            console.log(`Scheduler API response: ${response}`);
                            await context.sendActivity(response);
                            await next();
                            return;
                        }
                    } else {
                        console.log('CLU configuration incomplete. Skipping to QnA.');
                    }
                } catch (cluError) {
                    console.log(`Error occurred in CLU intent detection: ${cluError.message}`);
                    console.log('CLU unavailable. Skipping to QnA.');
                }

                const looksLikeAvailabilityRequest =
                    normalizedMessage.includes('available') ||
                    normalizedMessage.includes('availability') ||
                    normalizedMessage.includes('open slot') ||
                    normalizedMessage.includes('time slot');

                if (looksLikeAvailabilityRequest) {
                    console.log('Scheduler triggered');
                    console.log('Handling availability request after CLU did not return a scheduler intent.');
                    response = await this.scheduler.getAvailability();
                    console.log(`Scheduler API response: ${response}`);
                    await context.sendActivity(response);
                    await next();
                    return;
                }

                const looksLikeScheduleRequest =
                    normalizedMessage.includes('book') ||
                    normalizedMessage.includes('schedule') ||
                    normalizedMessage.includes('reserve');

                if (looksLikeScheduleRequest && normalizedMessage.includes('appointment')) {
                    console.log('Scheduler triggered');
                    console.log('Handling appointment request after CLU did not return a scheduler intent.');
                    response = await this.scheduler.scheduleAppointment(appointmentTime);
                    console.log(`Scheduler API response: ${response}`);
                    await context.sendActivity(response);
                    await next();
                    return;
                }

                // Step 2: Try Azure Custom Question Answering
                console.log('Attempting Azure Custom Question Answering...');
                const qnaConfig = this.QnAConfiguration;
                if (qnaConfig && qnaConfig.host && qnaConfig.endpointKey && qnaConfig.projectName && qnaConfig.deploymentName) {
                    try {
                        const qnaHost = qnaConfig.host.endsWith('/') ? qnaConfig.host : `${qnaConfig.host}/`;
                        const qnaUrl = `${qnaHost}language/:query-knowledgebases?projectName=${encodeURIComponent(qnaConfig.projectName)}&deploymentName=${encodeURIComponent(qnaConfig.deploymentName)}&api-version=2021-10-01`;

                        console.log(`User question: ${userMessage}`);

                        const qnaResponse = await fetch(qnaUrl, {
                            method: 'POST',
                            headers: {
                                'Content-Type': 'application/json',
                                'Ocp-Apim-Subscription-Key': qnaConfig.endpointKey
                            },
                            body: JSON.stringify({
                                question: userMessage
                            })
                        });

                        const qnaResult = await qnaResponse.json();
                        console.log(`QnA API response: ${JSON.stringify(qnaResult)}`);

                        const qnaAnswer = qnaResult.answers && qnaResult.answers[0] && qnaResult.answers[0].answer;
                        if (qnaAnswer && qnaAnswer !== 'No answer found') {
                            console.log('QnA triggered');
                            response = qnaAnswer;
                            console.log(`QnA Answer: ${response}`);
                            await context.sendActivity(response);
                            await next();
                            return;
                        }
                    } catch (qnaError) {
                        console.log(`Error occurred in Custom Question Answering: ${qnaError.message}`);
                        console.log('Custom Question Answering unavailable. Using fallback.');
                    }
                } else {
                    console.log('Custom Question Answering configuration incomplete. Using fallback.');
                }

                // Step 3: Fallback response if no other condition matches
                console.log('Fallback triggered');
                console.log('No scheduler or QnA result found. Using fallback response.');
                response = "I'm here to help with dental appointments and FAQs. You can ask me things like:\n" +
                    "• 'What appointments are available?'\n" +
                    "• 'Schedule an appointment for tomorrow'\n" +
                    "• 'Do you accept insurance?'\n" +
                    "• 'What are your office hours?'\n\n" +
                    "How can I assist you today?";

                await context.sendActivity(response);
                await next();
            } catch (error) {
                console.log(`Error occurred in onMessage handler: ${error.message}`);
                const errorResponse = "I encountered an error while processing your request. Please try again.";
                await context.sendActivity(errorResponse);
                await next();
            }
        });

        this.onMembersAdded(async (context, next) => {
        const membersAdded = context.activity.membersAdded;
        //write a custom greeting
        const welcomeText = '';
        for (let cnt = 0; cnt < membersAdded.length; ++cnt) {
            if (membersAdded[cnt].id !== context.activity.recipient.id) {
                await context.sendActivity(MessageFactory.text(welcomeText, welcomeText));
            }
        }
        // by calling next() you ensure that the next BotHandler is run.
        await next();
    });
    }
}

module.exports.DentaBot = DentaBot;
