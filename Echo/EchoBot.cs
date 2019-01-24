// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Teams;
using Microsoft.Bot.Connector.Teams;
using Microsoft.Bot.Schema;
using Microsoft.Bot.Schema.Teams;
using Microsoft.Extensions.Logging;

namespace Echo
{
    /// <summary>
    /// Represents a bot that processes incoming activities.
    /// For each user interaction, an instance of this class is created and the OnTurnAsync method is called.
    /// This is a Transient lifetime service.  Transient lifetime services are created
    /// each time they're requested. For each Activity received, a new instance of this
    /// class is created. Objects that are expensive to construct, or have a lifetime
    /// beyond the single turn, should be carefully managed.
    /// For example, the <see cref="MemoryStorage"/> object and associated
    /// <see cref="IStatePropertyAccessor{T}"/> object are created with a singleton lifetime.
    /// </summary>
    public class EchoBot : IBot
    {
        private readonly EchoAccessors _accessors;
        private readonly ILogger _logger;

        /// <summary>
        /// Initializes a new instance of the class.
        /// </summary>
        /// <param name="accessors">A class containing <see cref="IStatePropertyAccessor{T}"/> used to manage state.</param>
        /// <param name="loggerFactory">A <see cref="ILoggerFactory"/> that is hooked to the Azure App Service provider.</param>
        public EchoBot(EchoAccessors accessors, ILoggerFactory loggerFactory)
        {
            if (loggerFactory == null)
            {
                throw new System.ArgumentNullException(nameof(loggerFactory));
            }

            _logger = loggerFactory.CreateLogger<EchoBot>();
            _logger.LogTrace("Turn start.");
            _accessors = accessors ?? throw new System.ArgumentNullException(nameof(accessors));
        }

        /// <summary>
        /// Every conversation turn for our Echo Bot will call this method.
        /// There are no dialogs used, since it's "single turn" processing, meaning a single
        /// request and response.
        /// </summary>
        /// <param name="turnContext">A <see cref="ITurnContext"/> containing all the data needed
        /// for processing this conversation turn. </param>
        /// <param name="cancellationToken">(Optional) A <see cref="CancellationToken"/> that can be used by other objects
        /// or threads to receive notice of cancellation.</param>
        /// <returns>A <see cref="Task"/> that represents the work queued to execute.</returns>
        public async Task OnTurnAsync(ITurnContext turnContext, CancellationToken cancellationToken = default(CancellationToken))
        {
            // Handle Message activity type, which is the main activity type for shown within a conversational interface
            // Message activities may contain text, speech, interactive cards, and binary or unknown attachments.
            // see https://aka.ms/about-bot-activity-message to learn more about the message and other activity types
            if (turnContext.Activity.Type == ActivityTypes.Message)
            {

                // Before doing Teams specific stuff, get hold of the TeamsContext
                ITeamsContext teamsContext = turnContext.TurnState.Get<ITeamsContext>();

                // Now fetch the Team ID, Channel ID, and Tenant ID off of the incoming activity
                string incomingTeamId = teamsContext.Team.Id;
                string incomingChannelid = teamsContext.Channel.Id;
                string incomingTenantId = teamsContext.Tenant.Id;

                // Make an operation call to fetch the list of channels in the team, and print count of channels.
                ConversationList channels = await teamsContext.Operations.FetchChannelListAsync(incomingTeamId);
                await turnContext.SendActivityAsync($"You have {channels.Conversations.Count} channels in this team");

                // Make an operation call to fetch details of the team where the activity was posted, and print it.
                TeamDetails teamInfo = await teamsContext.Operations.FetchTeamDetailsAsync(incomingTeamId);
                await turnContext.SendActivityAsync($"Name of this team is {teamInfo.Name} and group-id is {teamInfo.AadGroupId}");
                
                // Get the conversation state from the turn context.
                CounterState state = await _accessors.CounterState.GetAsync(turnContext, () => new CounterState());

                // Bump the turn count for this conversation.
                state.TurnCount++;

                // Set the property using the accessor.
                await _accessors.CounterState.SetAsync(turnContext, state);

                // Save the new turn count into the conversation state.
                await _accessors.ConversationState.SaveChangesAsync(turnContext);

                // Echo back to the user whatever they typed.
                string responseMessage = $"Turn {state.TurnCount}: You sent '{turnContext.Activity.Text}'. Toodles!\n";
                await turnContext.SendActivityAsync(responseMessage);
            }
            else if (turnContext.Activity.Type == ActivityTypes.Invoke)
            {
                try
                {
                    MessagingExtensionResult composeExtensionResult = new MessagingExtensionResult
                    {
                        Type = "result",
                        AttachmentLayout = "list",
                        Attachments = new List<MessagingExtensionAttachment>(),
                    };

                    ThumbnailCard card = new ThumbnailCard
                    {
                        Title = "I'm a thumbnail",
                        Text = "Normally I'd be way more useful.",
                        Images = new CardImage[] { new CardImage("http://web.hku.hk/~jmwchiu/cats/cat10.jpg") },
                    };
                    composeExtensionResult.Attachments.Add(card.ToAttachment().ToMessagingExtensionAttachment());

                    InvokeResponse ir = new InvokeResponse
                    {
                        Body = new MessagingExtensionResponse
                        {
                            ComposeExtension = composeExtensionResult
                        },
                        Status = 200,
                    };

                    await turnContext.SendActivityAsync(new Activity
                    {
                        Value = ir,
                        Type = ActivityTypesEx.InvokeResponse,
                    });
                }
                catch (Exception ex)
                {
                    await turnContext.SendActivityAsync($"oops: {ex.InnerException}");
                }
            }
            else
            {
                await turnContext.SendActivityAsync($"{turnContext.Activity.Type} event detected");
            }
        }
    }
}
