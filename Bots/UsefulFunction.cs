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
using System.Linq;

namespace Microsoft.BotBuilderSamples.Bots
{
    public partial class TeamsConversationBot : TeamsActivityHandler
    {
        private async Task SearchCardActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken, string queryInfo)
        {
            var card = new HeroCard
            {
                Title = "Search Card",
                Text = "Search Results",
                Buttons = new List<CardAction>
                {
                    new CardAction
                    {
                        Type= ActionTypes.OpenUrl,
                        Title = "IntuneWiki",
                        Value = "https://intunewiki.com/index.php?search=" + queryInfo + "&title=Special%3ASearch&go=Go"
                    },
                    new CardAction
                    {
                        Type = ActionTypes.OpenUrl,
                        Title = "Google",
                        Value = "https://www.google.com/search?q=" + queryInfo
                    },
                    new CardAction
                    {
                        Type = ActionTypes.OpenUrl,
                        Title = "Code",
                        Value = "https://msazure.visualstudio.com/DefaultCollection/One/_search?type=code&text=" + queryInfo + "&filters=ProjectFilters%7BOne%7D&action=contents"
                    },
                    new CardAction
                    {
                        Type = ActionTypes.OpenUrl,
                        Title = "SharePoint",
                        Value = "https://microsoft.sharepoint.com/_layouts/15/search.aspx/?q=" + queryInfo
                    }
                }
            };
            await turnContext.SendActivityAsync(MessageFactory.Attachment(card.ToAttachment()));
        }

        private async Task MentionRollActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            var members = await TeamsInfo.GetMembersAsync(turnContext, cancellationToken);
            Random rd = new Random();

            var membersArray = members.ToArray();
            var tt = rd.Next(membersArray.Length);
            var randomMember = membersArray[tt];

            var mention = new Mention
            {
                Mentioned = randomMember,
                Text = $"<at>{XmlConvert.EncodeName(randomMember.Name)}</at>",
            };
            var replyActivity = MessageFactory.Text($"就决定是你了，去吧 {mention.Text}.");
            replyActivity.Entities = new List<Entity> { mention };

            await turnContext.SendActivityAsync(replyActivity, cancellationToken);
        }

        private async Task PairingProgrammingAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            var teamsChannelId = turnContext.Activity.TeamsGetChannelId();
            if (teamsChannelId != _MMDTeamGenaralChannelID)
            {
                await turnContext.SendActivityAsync("No permission for this command");
                return;
            }
            var members = await TeamsInfo.GetMembersAsync(turnContext, cancellationToken);
            var membersList = members.ToList();
            var mentionText = "Hello every, here is weekly pairing programing. Next is God's choice: \n\r";
            var mentionList = new List<Entity> { };
            Random rd = new Random();
            while (membersList.Count > 1)
            {
                var firstIndex = rd.Next(membersList.Count);
                var firstMember = membersList[firstIndex];
                var firstMention = new Mention
                {
                    Mentioned = firstMember,
                    Text = $"<at>{XmlConvert.EncodeName(firstMember.Name)}</at>",
                };
                membersList.RemoveAt(firstIndex);
                var secondIndex = rd.Next(membersList.Count);
                var secondMember = membersList[secondIndex];
                var secondMention = new Mention
                {
                    Mentioned = secondMember,
                    Text = $"<at>{XmlConvert.EncodeName(secondMember.Name)}</at>",
                };
                membersList.RemoveAt(secondIndex);
                mentionText += firstMention.Text + " & " + secondMention.Text + "\n\r";
                mentionList.Add(firstMention);
                mentionList.Add(secondMention);
            }

            if (membersList.Count == 1)
            {
                var lastMember = membersList[0];
                var lastMention = new Mention
                {
                    Mentioned = lastMember,
                    Text = $"<at>{XmlConvert.EncodeName(lastMember.Name)}</at>",
                };
                mentionText += "And " + lastMention.Text + ", feel free to join each group of them.";
                mentionList.Add(lastMention);
            }

            var replyActivity = MessageFactory.Text(mentionText);
            replyActivity.Entities = mentionList;

            await turnContext.SendActivityAsync(replyActivity, cancellationToken);
        }
    }
}