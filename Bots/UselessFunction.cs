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
        private async Task UpdateCardActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            var data = turnContext.Activity.Value as JObject;
            data["count"] = data["count"].Value<int>() + 1;
            data = JObject.FromObject(data);

            var card = new HeroCard
            {
                Title = "Welcome Card",
                Text = $"一顿操作猛如虎，一看输出 {data["count"].Value<int>()}.5",
                Buttons = new List<CardAction>
                        {
                            new CardAction
                            {
                                Type= ActionTypes.MessageBack,
                                Title = "Update Card",
                                Text = "UpdateCardAction",
                                Value = data
                            },
                            new CardAction
                            {
                                Type = ActionTypes.MessageBack,
                                Title = "Message all members",
                                Text = "MessageAllMembers"
                            },
                            new CardAction
                            {
                                Type = ActionTypes.MessageBack,
                                Title = "Delete card",
                                Text = "Delete"
                            }
                        }
            };

            var updatedActivity = MessageFactory.Attachment(card.ToAttachment());
            updatedActivity.Id = turnContext.Activity.ReplyToId;
            await turnContext.UpdateActivityAsync(updatedActivity, cancellationToken);
        }

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
        private async Task ShowWelcome(ITurnContext<IMessageActivity> turnContext)
        {
            var value = new JObject { { "count", 0 } };

            var card = new HeroCard
            {
                Title = "Welcome Card",
                Text = "Bonjour à tous, je suis Microsoft Fire.(点击此处进行翻译)",
                Buttons = new List<CardAction>
                        {
                            new CardAction
                            {
                                Type= ActionTypes.MessageBack,
                                Title = "Update Card",
                                Text = "UpdateCardAction",
                                Value = value
                            },
                            new CardAction
                            {
                                Type = ActionTypes.MessageBack,
                                Title = "Message all members",
                                Text = "MessageAllMembers"
                            }
                        }
            };

            await turnContext.SendActivityAsync(MessageFactory.Attachment(card.ToAttachment()));
        }
        private async Task DeleteCardActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            await turnContext.DeleteActivityAsync(turnContext.Activity.ReplyToId, cancellationToken);
        }

    }
}