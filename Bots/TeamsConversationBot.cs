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
        private string _MMDTeamGenaralChannelID;

        public TeamsConversationBot(IConfiguration config)
        {
            _appId = config["MicrosoftAppId"];
            _appPassword = config["MicrosoftAppPassword"];
            _authorPrincipalName = "zheta@microsoft.com";
            _MMDTeamGenaralChannelID = "19:fb37a43e619a4050bd722b299a308fe8@thread.tacv2";
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
                    await SendInfoToAllMemberAsync(turnContext, cancellationToken, sendInfo);
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
                case Constants.PairingProgramming.Name:
                case Constants.PairingProgramming.ShortName:
                    await PairingProgrammingAsync(turnContext, cancellationToken);
                    break;
                case Constants.NextMember.Name:
                case Constants.NextMember.ShortName:
                    await MentionNextMemberAsync(turnContext, cancellationToken);
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
                var mention = new Mention
                {
                    Mentioned = teamMember,
                    Text = $"<at>{XmlConvert.EncodeName(teamMember.Name)}</at>",
                };
                var replyActivity = MessageFactory.Text($"Welcome to the team {mention.Text}.");
                replyActivity.Entities = new List<Entity> { mention };
                await turnContext.SendActivityAsync(replyActivity, cancellationToken);
            }
        }



    }
}
