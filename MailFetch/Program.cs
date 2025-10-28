/*
 * Copyright 2025 LevelBlue
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

using CommandLine;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace MailFetch
{
    public class CLIOptions
    {
        [Option("app-id", HelpText = "Application Id", Required = true)]
        public string? AppId { get; set; }

        [Option("organization-id", HelpText = "Organization Id (Tenant Id)", Required = true)]
        public string? OrganizationId { get; set; }

        [Option("client-secret", HelpText = "Application Client Secret (password)", Required = true)]
        public string? ClientSecret { get; set; }

        [Option("username", HelpText = "User to collect email messages from", Required = true)]
        public string? Username { get; set; }

        private string? _output;
        [Option("output", HelpText = "Output Directory", Required = true)]
        public string? Output { get => _output; set { _output = Path.Combine(Directory.GetCurrentDirectory(), value); } }

        [Option("include-attachments", HelpText = "Download message attachments")]
        public bool IncludeAttachments { get; set; }

        [Option("all-results", HelpText = "Retrieve all results, not just the first page")]
        public bool AllResults { get; set; }

        public DateTime Started { get; } = DateTime.UtcNow;
    }
    class Program
    {
        private static readonly NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();
        public static Task<int> Main(string[] args)
        {
            var result = Parser.Default.ParseArguments<CLIOptions>(args);
            return result.MapResult(async (options) =>
            {
                try
                {
                    await Run(options);
                    return 0;
                }
                catch (Exception e)
                {
                    Console.Error.WriteLine(e);
                    Console.Error.WriteLine($"{e.GetType().Name}: {e.Message}");
                    return 2;
                }
            },
                errors => Task.FromResult(1));
        }

        private static async Task Run(CLIOptions options)
        {
            try
            {
                using var logConfig = new LogConfig(options);
                logger.Info("MailFetch started");
                var client = await CreateConfidentialClient(options);

                logger.Info("Reading mail for {Username:l}", options.Username);
                await ReadMail(options, client);
                logger.Debug("All done!");
            }
            catch (Exception e)
            {
                logger.Fatal(e, "Critical Error: ");
                throw;
            }
        }

        private static async Task ReadMail(CLIOptions options, Microsoft.Graph.GraphServiceClient client)
        {
            using var writer = new Writer(Path.Combine(options.Output, "results", $"results_{options.Started:yyyyMMddHHmmssfffffff}.json"));
            var messagesRequest = client.Users[options.Username!].Messages
                .Request();
            await messagesRequest.ForEach(client, await messagesRequest.GetAsync(), message => ProcessMessage(options, client, writer, message), options.AllResults);
        }

        private static async Task ProcessMessage(CLIOptions options, Microsoft.Graph.GraphServiceClient client, Writer writer, Microsoft.Graph.Message message)
        {
            var attachments = new List<Microsoft.Graph.Attachment>();
            if (options.IncludeAttachments)
            {
                var request = client.Users[options.Username!].Messages[message.Id].Attachments.Request();
                attachments = await request.ToList(client, await request.GetAsync(), options.AllResults);
            }
            writer.WriteLine(new { Message = message, Attachments = attachments });
        }

        private static async Task<Microsoft.Graph.GraphServiceClient> CreateConfidentialClient(CLIOptions options)
        {
            logger.Info("Using existing application credentials");
            return await CreateConfidentialClient(options.AppId!, options.OrganizationId!, options.ClientSecret!);
        }

        private static async Task<Microsoft.Graph.GraphServiceClient> CreateConfidentialClient(string appId, string organizationId, string clientSecret)
        {
            logger.Debug("AppId={AppId:l}, OrganizationId={OrganizationId:l}, ClientSecret={ClientSecret:l}", appId, organizationId, clientSecret);

            var app = ConfidentialClientApplicationBuilder
                .Create(appId)
                .WithTenantId(organizationId)
                .WithClientSecret(clientSecret)
                .Build();
            var authProvider = new ClientCredentialProvider(app);


            await app.AcquireTokenForClient(new string[] { "https://graph.microsoft.com/.default" }).ExecuteAsync();

            return new Microsoft.Graph.GraphServiceClient(authProvider);
        }
    }

    public static class JSONSerializer
    {
        private static readonly JsonSerializerSettings Settings = new JsonSerializerSettings
        {
            DateFormatString = "o",
            DateTimeZoneHandling = DateTimeZoneHandling.Utc,
            Converters = { new StringEnumConverter { } },
            ReferenceLoopHandling = ReferenceLoopHandling.Ignore
        };

        public static string SerializeObject(object? o)
        {
            return JsonConvert.SerializeObject(o, Settings);
        }
    }

    public sealed class Writer : IDisposable
    {
        private readonly StreamWriter _writer;
        public Writer(string filename)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(filename));
            _writer = new StreamWriter(filename);
        }

        public void Dispose()
        {
            _writer.Dispose();
        }

        public void WriteLine(object? o)
        {
            _writer.WriteLine(JSONSerializer.SerializeObject(o));
        }
    }

    public static class CollectionRequestExtensions
    {
        private static readonly NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();

        public static async Task Iterate<TRequest, TItem>(this TRequest request, Microsoft.Graph.IBaseClient client, Microsoft.Graph.ICollectionPage<TItem> page, Func<TItem, bool> processResult)
            where TRequest : Microsoft.Graph.IBaseRequest
        {
            request.Log("GET");
            var iterator = Microsoft.Graph.PageIterator<TItem>.CreatePageIterator(client, page, processResult);
            await iterator.IterateAsync();
        }

        public static Task Iterate<TRequest, TItem>(this TRequest request, Microsoft.Graph.IBaseClient client, Microsoft.Graph.ICollectionPage<TItem> page, Func<TItem, Task<bool>> processResult)
            where TRequest : Microsoft.Graph.IBaseRequest
        {
            return Iterate(request, client, page, (item) => processResult(item).GetAwaiter().GetResult());
        }

        public static async Task<List<TItem>> ToList<TRequest, TItem>(this TRequest request, Microsoft.Graph.IBaseClient client, Microsoft.Graph.ICollectionPage<TItem> page, bool all)
            where TRequest : Microsoft.Graph.IBaseRequest
        {
            List<TItem> results = new List<TItem>();
            bool ProcessResult(TItem item)
            {
                results.Add(item);
                return all; // Whether to keep processing
            }
            await Iterate(request, client, page, ProcessResult);
            return results;
        }

        public static async Task ForEach<TRequest, TItem>(this TRequest request, Microsoft.Graph.IBaseClient client, Microsoft.Graph.ICollectionPage<TItem> page, Action<TItem> action, bool all)
            where TRequest : Microsoft.Graph.IBaseRequest
        {
            bool ProcessResult(TItem item)
            {
                action(item);
                return all; // Whether to keep processing
            }
            await Iterate(request, client, page, ProcessResult);
        }

        public static Task ForEach<TRequest, TItem>(this TRequest request, Microsoft.Graph.IBaseClient client, Microsoft.Graph.ICollectionPage<TItem> page, Func<TItem, Task> action, bool all)
            where TRequest : Microsoft.Graph.IBaseRequest
        {
            async Task<bool> ProcessResult(TItem item)
            {
                await action(item);
                return all; // Whether to keep processing
            }
            return Iterate(request, client, page, ProcessResult);
        }

        public static TRequest Log<TRequest>(this TRequest request, string method)
            where TRequest : Microsoft.Graph.IBaseRequest
        {
            // request.Method is a bunch of nonsense, since the request is created before the method is actually specified
            var options = string.Join("&", request.QueryOptions.Select(option => $"{option.Name}={option.Value}"));
            logger.Debug("{Method:l} {Url:l}{Options:l}", method, request.RequestUrl, string.IsNullOrEmpty(options) ? "" : "?" + options);
            return request;
        }

        public static async Task<TItem?> FirstWhere<TRequest, TItem>(this TRequest request, Microsoft.Graph.IBaseClient client, Microsoft.Graph.ICollectionPage<TItem> page, Func<TItem, bool> predicate)
            where TRequest : Microsoft.Graph.IBaseRequest
            where TItem : class
        {
            // This method should only be used when the request doesn't support $filter
            TItem? match = null;
            bool ProcessResult(TItem item)
            {
                if (predicate(item))
                {
                    match = item;
                    return false;
                }
                return true;
            }
            await Iterate(request, client, page, ProcessResult);
            return match;
        }
    }
}
