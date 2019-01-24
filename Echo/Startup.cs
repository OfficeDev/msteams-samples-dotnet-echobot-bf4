﻿// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.BotFramework;
using Microsoft.Bot.Builder.Integration;
using Microsoft.Bot.Builder.Integration.AspNet.Core;
using Microsoft.Bot.Builder.Teams.Middlewares;
using Microsoft.Bot.Builder.Teams.StateStorage;
using Microsoft.Bot.Configuration;
using Microsoft.Bot.Connector.Authentication;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using System;
using System.Linq;

namespace Echo
{
    /// <summary>
    /// The Startup class configures services and the request pipeline.
    /// </summary>
    public class Startup
    {
        private ILoggerFactory _loggerFactory;
        private readonly bool _isProduction;

        public Startup(IHostingEnvironment env)
        {
            _isProduction = env.IsProduction();
            var builder = new ConfigurationBuilder()
                .SetBasePath(env.ContentRootPath)
                .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
                .AddJsonFile($"appsettings.{env.EnvironmentName}.json", optional: true)
                .AddEnvironmentVariables();

            Configuration = builder.Build();
        }

        /// <summary>
        /// Gets the configuration that represents a set of key/value application configuration properties.
        /// </summary>
        /// <value>
        /// The <see cref="IConfiguration"/> that represents a set of key/value application configuration properties.
        /// </value>
        public IConfiguration Configuration { get; }

        /// <summary>
        /// This method gets called by the runtime. Use this method to add services to the container.
        /// </summary>
        /// <param name="services">The <see cref="IServiceCollection"/> specifies the contract for a collection of service descriptors.</param>
        /// <seealso cref="IStatePropertyAccessor{T}"/>
        /// <seealso cref="https://docs.microsoft.com/en-us/aspnet/web-api/overview/advanced/dependency-injection"/>
        /// <seealso cref="https://docs.microsoft.com/en-us/azure/bot-service/bot-service-manage-channels?view=azure-bot-service-4.0"/>
        public void ConfigureServices(IServiceCollection services)
        {
            services.AddBot<EchoBot>(options =>
           {
               var secretKey = Configuration.GetSection("botFileSecret")?.Value;
               var botFilePath = Configuration.GetSection("botFilePath")?.Value;

               // Loads .bot configuration file and adds a singleton that your Bot can access through dependency injection.
               var botConfig = BotConfiguration.Load(botFilePath ?? @".\Echo.bot", secretKey);
               services.AddSingleton(sp => botConfig ?? throw new InvalidOperationException($"The .bot config file could not be loaded. ({botConfig})"));

               // Retrieve current endpoint.
               var environment = _isProduction ? "production" : "development";
               var service = botConfig.Services.FirstOrDefault(s => s.Type == "endpoint" && s.Name == environment);
               if (!(service is EndpointService endpointService))
               {
                   throw new InvalidOperationException($"The .bot file does not contain an endpoint with name '{environment}'.");
               }

               options.CredentialProvider = new ConfigurationCredentialProvider(this.Configuration);

               // Creates a logger for the application to use.
               ILogger logger = _loggerFactory.CreateLogger<EchoBot>();

               // Catches any errors that occur during a conversation turn and logs them.
               options.OnTurnError = async (context, exception) =>
              {
                  logger.LogError($"Exception caught : {exception}");
                  await context.SendActivityAsync("Sorry, it looks like something went wrong.");
              };

               // The Memory Storage used here is for local bot debugging only. When the bot
               // is restarted, everything stored in memory will be gone.
               IStorage dataStore = new MemoryStorage();

               // For production bots use the Azure Blob or
               // Azure CosmosDB storage providers. For the Azure
               // based storage providers, add the Microsoft.Bot.Builder.Azure
               // Nuget package to your solution. That package is found at:
               // https://www.nuget.org/packages/Microsoft.Bot.Builder.Azure/
               // Uncomment the following lines to use Azure Blob Storage
               // //Storage configuration name or ID from the .bot file.
               // const string StorageConfigurationId = "<STORAGE-NAME-OR-ID-FROM-BOT-FILE>";
               // var blobConfig = botConfig.FindServiceByNameOrId(StorageConfigurationId);
               // if (!(blobConfig is BlobStorageService blobStorageConfig))
               // {
               //    throw new InvalidOperationException($"The .bot file does not contain an blob storage with name '{StorageConfigurationId}'.");
               // }
               // // Default container name.
               // const string DefaultBotContainer = "<DEFAULT-CONTAINER>";
               // var storageContainer = string.IsNullOrWhiteSpace(blobStorageConfig.Container) ? DefaultBotContainer : blobStorageConfig.Container;
               // IStorage dataStore = new Microsoft.Bot.Builder.Azure.AzureBlobStorage(blobStorageConfig.ConnectionString, storageContainer);

               // Create Conversation State object.
               // The Conversation State object is where we persist anything at the conversation-scope.
               var conversationState = new TeamSpecificConversationState(dataStore);
               options.Middleware.Add(new DropNonTeamsActivitiesMiddleware());
               options.Middleware.Add(new TeamsMiddleware(
                  new ConfigurationCredentialProvider(this.Configuration)));

               options.State.Add(conversationState);
           });

            // Create and register state accessors.
            // Accessors created here are passed into the IBot-derived class on every turn.
            services.AddSingleton<EchoAccessors>(sp =>
           {
               var options = sp.GetRequiredService<IOptions<BotFrameworkOptions>>().Value;
               if (options == null)
               {
                   throw new InvalidOperationException("BotFrameworkOptions must be configured prior to setting up the state accessors");
               }

               var conversationState = options.State.OfType<TeamSpecificConversationState>().FirstOrDefault();
               if (conversationState == null)
               {
                   throw new InvalidOperationException("ConversationState must be defined and added before adding conversation-scoped state accessors.");
               }

               // Create the custom state accessor.
               // State accessors enable other components to read and write individual properties of state.
               var accessors = new EchoAccessors(conversationState)
               {
                   CounterState = conversationState.CreateProperty<CounterState>(EchoAccessors.CounterStateName),
               };

               return accessors;
           });
        }

        public void Configure(IApplicationBuilder app, IHostingEnvironment env, ILoggerFactory loggerFactory)
        {
            _loggerFactory = loggerFactory;

            app.UseDefaultFiles()
                .UseStaticFiles()
                .UseBotFramework();
        }
    }
}
