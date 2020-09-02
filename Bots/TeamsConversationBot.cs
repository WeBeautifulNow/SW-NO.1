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
            var input = turnContext.Activity.Text.Trim();
            var inputUpper = input.ToUpper();
            if (inputUpper.StartsWith(Constants.Search.Name.ToUpper())) {
                var searchParams = input.Remove(0, input.IndexOf("+") + 1);
                await SearchCardActivityAsync(turnContext, cancellationToken, searchParams);
            } else if (inputUpper.StartsWith(Constants.SendMessageToAll.Name.ToUpper())) {
                var sendInfo = input.Remove(0, input.IndexOf("+") + 1);
                await SendInfoToAllMemberAsync(turnContext, cancellationToken, sendInfo);
            } else if (string.Equals(input, Constants.ShowAllCommands.Name, StringComparison.InvariantCultureIgnoreCase)
                || string.Equals(input, Constants.ShowAllCommands.ShortName, StringComparison.InvariantCultureIgnoreCase)) {
                ShowHelpInfo(ref reply);
                await turnContext.SendActivityAsync(reply, cancellationToken: cancellationToken);
            } else if (string.Equals(input, Constants.MentionMe.Name, StringComparison.InvariantCultureIgnoreCase)
                || string.Equals(input, Constants.MentionMe.ShortName, StringComparison.InvariantCultureIgnoreCase)) {
                await MentionActivityAsync(turnContext, cancellationToken);
            } else if (string.Equals(input, Constants.Roll.Name, StringComparison.InvariantCultureIgnoreCase)
                || string.Equals(input, Constants.Roll.ShortName, StringComparison.InvariantCultureIgnoreCase)) {
                await MentionRollActivityAsync(turnContext, cancellationToken);
            } else if (string.Equals(input, Constants.PairingProgramming.Name, StringComparison.InvariantCultureIgnoreCase)
                || string.Equals(input, Constants.PairingProgramming.ShortName, StringComparison.InvariantCultureIgnoreCase)) {
                await PairingProgrammingAsync(turnContext, cancellationToken);
            } else if (string.Equals(input, Constants.NextMember.Name, StringComparison.InvariantCultureIgnoreCase)
              || string.Equals(input, Constants.NextMember.ShortName, StringComparison.InvariantCultureIgnoreCase)) {
                await MentionNextMemberAsync(turnContext, cancellationToken);
            } else if (string.Equals(input, Constants.Thanks.Name, StringComparison.InvariantCultureIgnoreCase)) {
                Thanks(ref reply);
                await turnContext.SendActivityAsync(reply, cancellationToken: cancellationToken);
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
