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
    }
}