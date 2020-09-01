// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System;
using System.Text.RegularExpressions;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Teams;
using Microsoft.Bot.Connector.Authentication;
using Microsoft.Bot.Schema;
using Microsoft.Bot.Schema.Teams;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json.Linq;

namespace Microsoft.BotBuilderSamples.Bots
{
    public partial class TeamsConversationBot : TeamsActivityHandler
    {
        private string _appId;
        private string _appPassword;
        private string _authorPrincipalName;

        public TeamsConversationBot(IConfiguration config)
        {
            _appId = config["MicrosoftAppId"];
            _appPassword = config["MicrosoftAppPassword"];
            _authorPrincipalName = "zheta@microsoft.com";
        }

        protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            turnContext.Activity.RemoveRecipientMention();

            var reply = "Beyond the scope of my capabilities, please contact the developer to let me upgrade";

            switch (turnContext.Activity.Text.Trim())
            {
                case var someVal when someVal.StartsWith(Constants.Search.Name):
                    var searchParams = someVal.Remove(0, someVal.IndexOf("+") + 1);
                    await SearchCardActivityAsync(turnContext, cancellationToken, searchParams);
                    break;
                case var someVal when someVal.StartsWith(Constants.SendMessageToAll.Name):
                    var sendInfo = someVal.Remove(0, someVal.IndexOf("+") + 1);
                    await SendInfoToAllMember(turnContext, cancellationToken, sendInfo);
                    break;
                case Constants.ShowAllCommands.Name:
                case Constants.ShowAllCommands.ShortName:
                    ShowHelpInfo(ref reply);
                    await turnContext.SendActivityAsync(reply, cancellationToken: cancellationToken);
                    break;
                case Constants.MentionMe.Name:
                case Constants.MentionMe.ShortName:
                    await MentionActivityAsync(turnContext, cancellationToken);
                    break;
                case Constants.Roll.Name:
                case Constants.Roll.ShortName:
                    await MentionRollActivityAsync(turnContext, cancellationToken);
                    break;

                default:
                    await turnContext.SendActivityAsync(reply, cancellationToken: cancellationToken);
                    break;

            }
        }

        protected override async Task OnTeamsMembersAddedAsync(IList<TeamsChannelAccount> membersAdded, TeamInfo teamInfo, ITurnContext<IConversationUpdateActivity> turnContext, CancellationToken cancellationToken)
        {
            foreach (var teamMember in membersAdded)
            {
                await turnContext.SendActivityAsync(MessageFactory.Text($"Welcome to the team {teamMember.GivenName} {teamMember.Surname}."), cancellationToken);
            }
        }



    }
}
