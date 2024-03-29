﻿using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using Microsoft.Graph;
using Microsoft.Extensions.Configuration;

namespace MSGraphTest
{
    class Program
    {
        private static GraphServiceClient _graphServiceClient;
        private static HttpClient _httpClient;

        private static IConfigurationRoot LoadAppSettings()
        {
            try
                {
                var config = new ConfigurationBuilder()
                .SetBasePath(System.IO.Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", false, true)
                .Build();

                // Validate required settings
                if (string.IsNullOrEmpty(config["applicationId"]) || string.IsNullOrEmpty(config["applicationSecret"]) || string.IsNullOrEmpty(config["redirectUri"]) || string.IsNullOrEmpty(config["tenantId"]))
                {
                    return null;
                }

                return config;
                }
            catch (System.IO.FileNotFoundException)
            {
            return null;
            }
        }

        private static IAuthenticationProvider CreateAuthorizationProvider(IConfigurationRoot config)
        {
            var clientId = config["applicationId"];
            var clientSecret = config["applicationSecret"];
            var redirectUri = config["redirectUri"];
            var authority = $"https://login.microsoftonline.com/{config["tenantId"]}/v2.0";

            List<string> scopes = new List<string>();
            scopes.Add("https://graph.microsoft.com/.default");

            var cca = new ConfidentialClientApplication(clientId, authority, redirectUri, new ClientCredential(clientSecret), null, null);
            return new MsalAuthenticationProvider(cca, scopes.ToArray());
        }

        private static GraphServiceClient GetAuthenticatedGraphClient(IConfigurationRoot config)
        {
            var authenticationProvider = CreateAuthorizationProvider(config);
            _graphServiceClient = new GraphServiceClient(authenticationProvider);
            return _graphServiceClient;
        }

        private static HttpClient GetAuthenticatedHTTPClient(IConfigurationRoot config)
        {
            var authenticationProvider = CreateAuthorizationProvider(config);
            _httpClient = new HttpClient(new AuthHandler(authenticationProvider, new HttpClientHandler()));
            return _httpClient;
        }


        static void Main(string[] args)
        {
            var config = LoadAppSettings();
            if (null == config)
            {
            Console.WriteLine("Missing or invalid appsettings.json file. Please see README.md for configuration instructions.");
            return;
            }

            //Query using Graph SDK (preferred when possible)
            GraphServiceClient graphClient = GetAuthenticatedGraphClient(config);
            List<QueryOption> options = new List<QueryOption>
            {
                new QueryOption("$top", "10")
            };
            int pageCount = 0;
            var graphResult = graphClient.Users.Request().GetAsync().Result;
            int userCount = 0;
            while (true)
            {
                pageCount++;
                Console.WriteLine("Graph SDK Result Page: {0}",pageCount);

                foreach (var user in graphResult){
                    userCount++;
                    Console.WriteLine("User #{0}: " + user.DisplayName,userCount);
                }

                if (graphResult.NextPageRequest != null)
                {
                    graphResult = graphResult.NextPageRequest.GetAsync().Result;
                }
                else
                {
                    break;
                }

            }

            
        /*
            //Direct query using HTTPClient (for beta endpoint calls or not available in Graph SDK)
            HttpClient httpClient = GetAuthenticatedHTTPClient(config);
            Uri Uri = new Uri("https://graph.microsoft.com/v1.0/users?$top=1");
            //Uri Uri = new Uri("https://graph.microsoft.com/v1.0/users");
            
            var httpResult = httpClient.GetStringAsync(Uri).Result;

            Console.WriteLine("HTTP Result");
            Console.WriteLine(httpResult);       
        
         */



        }
    }
}
