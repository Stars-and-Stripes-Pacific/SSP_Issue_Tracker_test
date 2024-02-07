const { TeamsActivityHandler, CardFactory, TurnContext } = require("botbuilder");
//var spauth = require('node-sp-auth');
//var request = require('request-promise');
//var $REST = require("gd-sprest");
const https = require('https');

class TeamsBot extends TeamsActivityHandler {
  constructor() {
    super();

    this.onMessage(async (context, next) => {
      console.log("Running with Message Activity.");
      const removedMentionText = TurnContext.removeRecipientMention(context.activity);
      const txt = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim();

      const options = {
        protocol: 'https:',
        host: '0b16-202-212-180-65.ngrok-free.app',
        path: '/api',
        method: 'GET',
      };
      
      var categories = '';
      const req = https.request(options, (res) => {
          res.on('data', (chunk) => {
              console.log(`BODY: ${chunk}`);
              //context.sendActivity('Echo: start');
              categories = chunk;
          });
          res.on('end', () => {
              console.log('No more data in response.');
              //await context.sendActivity(`Echo: ${txt}`);
              //context.sendActivity(`Echo: ${chunk}`);
              // By calling next() you ensure that the next BotHandler is run.
              //next();

          });
      })
      
      req.on('error', (e) => {
        console.error(`problem with request: ${e.message}`);
      });
      
      req.end();

      //await context.sendActivity(`Echo: ${txt}`);
      await context.sendActivity(`Echo: ${categories}`);
      // By calling next() you ensure that the next BotHandler is run.
      await next();
    });

    // Listen to MembersAdded event, view https://docs.microsoft.com/en-us/microsoftteams/platform/resources/bot-v3/bots-notifications for more events
    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      for (let cnt = 0; cnt < membersAdded.length; cnt++) {
        if (membersAdded[cnt].id) {
          await context.sendActivity(
            `Hi there! I'm a Teams bot that will echo what you said to me.`
          );
          break;
        }
      }
      await next();
    });
  }
}

module.exports.TeamsBot = TeamsBot;
