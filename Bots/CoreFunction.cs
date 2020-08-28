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
        static void ShowHelpInfo(ref string reply)
        {
            reply = "search+: intunewiki,google,devops搜索，在机器人中以search+开头即可，比如 search+kusto" + "\n\r" +
                "all+: 给所在team或者群聊中的每个人发一条私信，比如 all+大家好" + "\n\r" +
                "roll: 在所在team或者群聊中随机抽中一个强者,比如 roll" + "\n\r";

        }
        private async Task SendInfoToAllMember(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken, string sendInfo)
        {
            var teamsChannelId = turnContext.Activity.TeamsGetChannelId();
            if (teamsChannelId is null)
            {
                await turnContext.SendActivityAsync(MessageFactory.Text("This feature can only used in groups."), cancellationToken);
                return;
            }
            var teamInfo = turnContext.Activity.TeamsGetTeamInfo();
            var userName = turnContext.Activity.From.Name;
            var serviceUrl = turnContext.Activity.ServiceUrl;
            var credentials = new MicrosoftAppCredentials(_appId, _appPassword);
            ConversationReference conversationReference = null;

            var members = await TeamsInfo.GetMembersAsync(turnContext, cancellationToken);

            foreach (var teamMember in members)
            {
                var memberName = teamMember.UserPrincipalName;
                var proactiveCard = new HeroCard
                {
                    Title = "Message from " + userName + " in group " + teamInfo.Name,
                    Text = sendInfo,
                    Buttons = new List<CardAction>
                    {
                        new CardAction
                            {
                                Type = ActionTypes.OpenUrl,
                                Title = "Chat with " + userName,
                                Value = "https://teams.microsoft.com/l/chat/0/0?users=" + memberName + "&message=Hi%20there%20"
                            },
                        new CardAction
                            {
                                Type = ActionTypes.OpenUrl,
                                Title = "Contact the developer",
                                Text = "Any suggestions, feel free to contact with developer >^<",
                                Value = "https://teams.microsoft.com/l/chat/0/0?users=" + _authorPrincipalName + "&message=Hi%20there%20"
                            },
                    }
                };

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
                                await t2.SendActivityAsync(MessageFactory.Attachment(proactiveCard.ToAttachment()), c2);
                            },
                            cancellationToken);
                    },
                    cancellationToken);
            }
            await turnContext.SendActivityAsync(MessageFactory.Text("All messages have been sent."), cancellationToken);
        }
    }
}