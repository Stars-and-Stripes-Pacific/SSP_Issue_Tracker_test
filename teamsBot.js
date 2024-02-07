const { TeamsActivityHandler, CardFactory, TurnContext } = require("botbuilder");
var spauth = require('node-sp-auth');
var request = require('request-promise');
var $REST = require("gd-sprest");

class TeamsBot extends TeamsActivityHandler {
  constructor() {
    super();

    this.onMessage(async (context, next) => {
      console.log("Running with Message Activity.");
      const removedMentionText = TurnContext.removeRecipientMention(context.activity);
      const txt = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim();
      var url = "https://stripes.sharepoint.com/sites/SSPDIDevNS-Noriko--TestIssueTracker";
      // SharePoint access
      spauth.getAuth(url, {
          username: "sekiya.takeichiro@stripes.com",
          password: "1live@srealman",
          online: true
      }).then(options => {
          // Log
          console.log("Connected to SPO");
      
          // Code Continues in 'Generate the Request'
          // Get the web
          var info = $REST.Web(url)
              .Lists("Category")
              .getInfo();
      
          for (var key in options.headers) {
              // Set the header
              info.headers[key] = options.headers[key];
          }
          var categorystr = '';
          // Execute the request, based on the method
          request[info.method == "GET" ? "get" : "post"]({
              headers: info.headers,
              //url: info.url,
              url: info.url + "/items",
              body: info.data
          }).then(
              // Success
              response => {
                  //console.log(response);
      
                  var obj = JSON.parse(response).d;
                  //console.log(obj);
                  if (obj.results && obj.results.length > 0) {
                      // Parse the results
                      for (var i = 0; i < obj.results.length; i++) {
                          // Log
                          //console.log(obj.results[i]);
                          console.log(obj.results[i]['Title']);
                          categorystr += obj.results[i]['Title']
                      }
                  } else {
                  }
              },
              // Error
              error => {
              }
          );
      });

      //await context.sendActivity(`Echo: ${txt}`);
      await context.sendActivity(`Echo: ${categorystr}`);
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
