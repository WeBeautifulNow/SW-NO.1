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
            reply = "Here are all commands(some need permission): \n\r"
                    + Constants.Search.Name + "${search content}\n\r"
                    + Constants.SendMessageToAll.Name + "${message}, for example: all+hello every\n\r"
                    + Constants.ShowAllCommands.Name + " or " + Constants.ShowAllCommands.ShortName + "\n\r"
                    + Constants.MentionMe.Name + " or " + Constants.MentionMe.ShortName + "\n\r"
                    + Constants.Roll.Name + " or " + Constants.Roll.ShortName + "\n\r"
                    + Constants.PairingProgramming.Name + " or " + Constants.PairingProgramming.ShortName + "\n\r"
                    + Constants.NextMember.Name + " or " + Constants.NextMember.ShortName + "\n\r";
        }
        private async Task SendInfoToAllMemberAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken, string sendInfo)
        {
            var teamsChannelId = turnContext.Activity.TeamsGetChannelId();
            if (teamsChannelId is null)
            {
                await turnContext.SendActivityAsync(MessageFactory.Text("This feature can only used in team."), cancellationToken);
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
                    Title = "Message from " + userName + " in team " + teamInfo.Name,
                    Text = sendInfo,
                    Buttons = new List<CardAction>
                    {
                        new CardAction
                            {
                                Type = ActionTypes.OpenUrl,
                                Title = "Chat with " + userName,
                                Value = "https://teams.microsoft.com/l/chat/0/0?users=" + _authorPrincipalName + "&message=Hi%20there%20"
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
        private async Task MentionNextMemberAsync(ITurnContext<IMessageActivity> turnContext, CancellationToken cancellationToken)
        {
            var members = await TeamsInfo.GetMembersAsync(turnContext, cancellationToken);
            var iterator = members.GetEnumerator();
            var operatorName = turnContext.Activity.From.Name;
            // nextMember is the preview element in iterator
            var preMember = iterator.Current;

            while (iterator.MoveNext())
            {
                TeamsChannelAccount current = iterator.Current;
                if (current.Name == operatorName)
                {
                    break;
                }
                preMember = iterator.Current;
            }
            // If preMember is null, that means operator is the first element in iterator. So we need return the last element
            if (preMember == null)
            {
                preMember = iterator.Current;
                while (iterator.MoveNext())
                {
                    preMember = iterator.Current;
                }
            }

            var mention = new Mention
            {
                Mentioned = preMember,
                Text = $"<at>{XmlConvert.EncodeName(preMember.Name)}</at>",
            };
            var replyActivity = MessageFactory.Text($"It's your turn {mention.Text}.");
            replyActivity.Entities = new List<Entity> { mention };

            await turnContext.SendActivityAsync(replyActivity, cancellationToken);
        }

    }
}