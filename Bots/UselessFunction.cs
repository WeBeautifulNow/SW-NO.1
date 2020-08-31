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
        private async Task MentionActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            var mention = new Mention
            {
                Mentioned = turnContext.Activity.From,
                Text = $"<at>{XmlConvert.EncodeName(turnContext.Activity.From.Name)}</at>",
            };

            var replyActivity = MessageFactory.Text($"Hello {mention.Text}.");
            replyActivity.Entities = new List<Entity> { mention };

            await turnContext.SendActivityAsync(replyActivity, cancellationToken);
        }
    }
}