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
    public class TeamsConversationBot : TeamsActivityHandler
    {
        private string _appId;
        private string _appPassword;

        public TeamsConversationBot(IConfiguration config)
        {
            _appId = config["MicrosoftAppId"];
            _appPassword = config["MicrosoftAppPassword"];
        }

        protected override async Task OnMessageActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            turnContext.Activity.RemoveRecipientMention();

            var reply = "Beyond the scope of my capabilities, please contact the developer to let me upgrade";

            switch (turnContext.Activity.Text.Trim())
            {
                case var someVal when someVal.StartsWith("search+"):
                    var searchParams = someVal.Remove(0, someVal.IndexOf("+") + 1);
                    await SearchCardActivityAsync(turnContext, cancellationToken, searchParams);
                    break;
                case var someVal when someVal.StartsWith("all+"):
                    var sendInfo = someVal.Remove(0, someVal.IndexOf("+") + 1);
                    await SendInfoToAllMember(turnContext, cancellationToken, sendInfo);
                    break;
                case "help":
                    ShowHelpInfo(ref reply);
                    await turnContext.SendActivityAsync(reply, cancellationToken: cancellationToken);
                    break;
                case "MentionMe":
                    await MentionActivityAsync(turnContext, cancellationToken);
                    break;
                case "roll":
                    await MentionRollActivityAsync(turnContext, cancellationToken);
                    break;
                case "UpdateCardAction":
                    await UpdateCardActivityAsync(turnContext, cancellationToken);
                    break;
                case "Delete":
                    await DeleteCardActivityAsync(turnContext, cancellationToken);
                    break;
                case "MessageAllMembers":
                    await MessageAllMembersAsync(turnContext, cancellationToken);
                    break;
                case "Show Welcome":
                    await ShowWelcome(turnContext);
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

        private async Task DeleteCardActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            await turnContext.DeleteActivityAsync(turnContext.Activity.ReplyToId, cancellationToken);
        }

        private async Task MessageAllMembersAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            var teamsChannelId = turnContext.Activity.TeamsGetChannelId();
            var serviceUrl = turnContext.Activity.ServiceUrl;
            var credentials = new MicrosoftAppCredentials(_appId, _appPassword);
            ConversationReference conversationReference = null;

            var members = await TeamsInfo.GetMembersAsync(turnContext, cancellationToken);

            foreach (var teamMember in members)
            {
                var proactiveMessage = MessageFactory.Text($"Hello {teamMember.GivenName} {teamMember.Surname}. I'm a Teams conversation bot.");

                var conversationParameters = new ConversationParameters
                {
                    IsGroup = false,
                    Bot = turnContext.Activity.Recipient,
                    Members = new ChannelAccount[] { teamMember },
                    TenantId = turnContext.Activity.Conversation.TenantId,
                };

                await ((BotFrameworkAdapter)turnContext.Adapter).CreateConversationAsync(
                    teamsChannelId,
                    serviceUrl,
                    credentials,
                    conversationParameters,
                    async (t1, c1) =>
                    {
                        conversationReference = t1.Activity.GetConversationReference();
                        await ((BotFrameworkAdapter)turnContext.Adapter).ContinueConversationAsync(
                            _appId,
                            conversationReference,
                            async (t2, c2) =>
                            {
                                await t2.SendActivityAsync(proactiveMessage, c2);
                            },
                            cancellationToken);
                    },
                    cancellationToken);
            }

            await turnContext.SendActivityAsync(MessageFactory.Text("All messages have been sent."), cancellationToken);
        }

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

        private async Task MentionRollActivityAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            var members = await TeamsInfo.GetMembersAsync(turnContext, cancellationToken);
            Random rd = new Random();
            var iterator = members.GetEnumerator();
            int count = 1;
            TeamsChannelAccount current = iterator.Current;
            while (iterator.MoveNext())
            {
                count++;
                if (rd.Next(count) == 0)
                {
                    current = iterator.Current;
                }
            }

            var mention = new Mention
            {
                Mentioned = current,
                Text = $"<at>{XmlConvert.EncodeName(current.Name)}</at>",
            };
            var prefix = "";
            switch (current.Name)
            {
                case "Croff Zhong":
                case "Aaron Yu":
                case "Youle Chen":
                case "Lingxiao Hang":
                    prefix = "貌若潘安";
                    break;
                case "Lijun Ma":
                case "Liu He":
                case "Ashley Yang":
                case "Daisy Zhao":
                case "Jingjing Han":
                    prefix = "倾国倾城";
                    break;
                case "Selin Luo":
                    prefix = "髣髴兮若轻云之蔽月，飘飖兮若流风之回雪";
                    break;
                default:
                    break;
            }
            var replyActivity = MessageFactory.Text($"就决定是你了，去吧 {prefix}{mention.Text}.");
            replyActivity.Entities = new List<Entity> { mention };

            await turnContext.SendActivityAsync(replyActivity, cancellationToken);
        }

        static void ShowHelpInfo(ref string reply)
        {
            reply = "search+: intunewiki,google,devops搜索，在机器人中以search+开头即可，比如 search+kusto" + "\n\r" +
                "all+: 给所在team或者群聊中的每个人发一条私信，比如 all+大家好" + "\n\r" +
                "roll: 在所在team或者群聊中随机抽中一个强者,比如 roll" + "\n\r";

        }

        private async Task SendInfoToAllMember(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken, string sendInfo)
        {
            var teamsChannelId = turnContext.Activity.TeamsGetChannelId();
            var serviceUrl = turnContext.Activity.ServiceUrl;
            var credentials = new MicrosoftAppCredentials(_appId, _appPassword);
            ConversationReference conversationReference = null;

            var members = await TeamsInfo.GetMembersAsync(turnContext, cancellationToken);

            foreach (var teamMember in members)
            {
                var proactiveMessage = MessageFactory.Text(sendInfo);

                var conversationParameters = new ConversationParameters
                {
                    IsGroup = false,
                    Bot = turnContext.Activity.Recipient,
                    Members = new ChannelAccount[] { teamMember },
                    TenantId = turnContext.Activity.Conversation.TenantId,
                };

                await ((BotFrameworkAdapter)turnContext.Adapter).CreateConversationAsync(
                    teamsChannelId,
                    serviceUrl,
                    credentials,
                    conversationParameters,
                    async (t1, c1) =>
                    {
                        conversationReference = t1.Activity.GetConversationReference();
                        await ((BotFrameworkAdapter)turnContext.Adapter).ContinueConversationAsync(
                            _appId,
                            conversationReference,
                            async (t2, c2) =>
                            {
                                await t2.SendActivityAsync(proactiveMessage, c2);
                            },
                            cancellationToken);
                    },
                    cancellationToken);
            }

            await turnContext.SendActivityAsync(MessageFactory.Text("All messages have been sent."), cancellationToken);
        }
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
                    }
                }
            };

            await turnContext.SendActivityAsync(MessageFactory.Attachment(card.ToAttachment()));

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
    }
}
