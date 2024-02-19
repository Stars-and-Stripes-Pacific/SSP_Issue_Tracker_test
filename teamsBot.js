const { TeamsActivityHandler, CardFactory, TurnContext } = require("botbuilder");
const https = require('https');

const state = {
  init:     0,
  category: 1,
  issue:    2,
}

// URL path definition
const urladdr      = 'cee3-202-212-180-65.ngrok-free.app'
const categorypath = '/api/category';
const issuepath    = '/api/issue?key=';

class TeamsBot extends TeamsActivityHandler {
  constructor() {
    super();

    this.currentstate = state.init;
    //var categorynum;
    this.categorytext = '';

    this.onMessage(async (context, next) => {
      console.log("Running with Message Activity.");
      const removedMentionText = TurnContext.removeRecipientMention(context.activity);
      const txt = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim();

      console.log("state:" + this.currentstate );
      // initial
      if(this.currentstate == state.init){
        const options = {
          protocol: 'https:',
          host: urladdr,
          path: categorypath,
          method: 'GET',
        };
        
        var categories = '';
        const req = https.request(options, (res) => {
            res.on('data', (chunk) => {
                console.log(`BODY: ${chunk}`);
                //context.sendActivity('Echo: start');
                categories = chunk;
                //await context.sendActivity(`Echo: ${txt}`);
                context.sendActivity(`Echo: ${categories}`);
            });
            res.on('end', () => {
                console.log('No more data in response.');
                //await context.sendActivity(`Echo: ${txt}`);
                //context.sendActivity(`Echo: ${chunk}`);
                // By calling next() you ensure that the next BotHandler is run.
                //next();
                this.currentstate = state.category;  
            });
        })
        
        req.on('error', (e) => {
          console.error(`problem with request: ${e.message}`);
        });
        
        req.end();

      // category
      }else if(this.currentstate == state.category){
        // set user input to variable
        this.categorytext = txt;
        this.currentstate = state.issue;
      }else if(this.currentstate == state.issue){
        const options = {
          protocol: 'https:',
          host: urladdr,
          path: issuepath + this.categorytext,
          method: 'GET',
        };
        
        var categories = '';
        const req = https.request(options, (res) => {
            res.on('data', (chunk) => {
                console.log(`BODY: ${chunk}`);
                //context.sendActivity('Echo: start');
                categories = chunk;
                //await context.sendActivity(`Echo: ${txt}`);
                context.sendActivity(`Echo: ${categories}`);
            });
            res.on('end', () => {
                console.log('No more data in response.');
                //await context.sendActivity(`Echo: ${txt}`);
                //context.sendActivity(`Echo: ${chunk}`);
                // By calling next() you ensure that the next BotHandler is run.
                //next();
                this.currentstate = state.init;  
            });
        })
        
        req.on('error', (e) => {
          console.error(`problem with request: ${e.message}`);
        });
        
        req.end();
        //this.currentstate = state.init;  
      }else{
        console.log("error state mismatch");
      }

      await context.sendActivity(`Echo: ${txt}`);
      //await context.sendActivity(`Echo: ${categories}`);
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
