using System.Text.Json;
using System.Net.Http.Headers;

using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Client;


namespace app
{
    public class WebhookProxy
    {
        private readonly ILogger<WebhookProxy> _logger;
        private readonly HttpClient _httpClient;
        private readonly string TENANT_ID = Environment.GetEnvironmentVariable("tenantId");
        private readonly string CLIENT_ID = Environment.GetEnvironmentVariable("clientId");
        private readonly string CLIENT_SECRET = Environment.GetEnvironmentVariable("clientSecret");
        private readonly string GRAPH_API_SCOPE = "https://graph.microsoft.com/.default";
        private readonly string GRAPH_API_URL = "https://graph.microsoft.com/v1.0/teams/{team-id}/channels/{channel-id}/messages";

        public WebhookProxy(HttpClient httpClient, ILogger<WebhookProxy> logger)
        {
            _httpClient = httpClient;
            _logger = logger;
        }

        [Function("WebhookProxy")]
        public async Task<HttpResponseData> Run([HttpTrigger(AuthorizationLevel.Function, "post")] HttpRequestData req)
        {
            _logger.LogInformation("WebhookProxy function processed a request.");


            // Step 1: Extract 'signature'
            var query = System.Web.HttpUtility.ParseQueryString(req.Url.Query);
            var signature = query["signature"];

            if (string.IsNullOrEmpty(signature))
            {
                _logger.LogError("No signature found in the request URL.");
                var errorResponse = req.CreateResponse(System.Net.HttpStatusCode.BadRequest);
                await errorResponse.WriteStringAsync("Signature parameter is missing");
                return errorResponse;
            }

            _logger.LogInformation($"Signature parameter: {signature}");


            // Step 2: Look up the team and channel corresponding to the signature
            var lookupResult = LookupTeamAndChannelBySignature(signature);
            if (lookupResult.Item1 == null || lookupResult.Item2 == null)
            {
                _logger.LogError($"No team or channel found for signature: {signature}");
                var errorResponse = req.CreateResponse(System.Net.HttpStatusCode.BadRequest);
                await errorResponse.WriteStringAsync("Invalid signature");
                return errorResponse;
            }

            var teamId = lookupResult.Item1;
            var channelId = lookupResult.Item2;

            // Step 3: Read the request body (this would be the message content)
            var requestBody = await req.ReadAsStringAsync();
            using var jsonDoc = JsonDocument.Parse(requestBody);
            var messageContent = jsonDoc.RootElement.GetProperty("message").GetString();

            // Step 4: Send the message to the corresponding Teams channel via Graph API
            // var accessToken = await GetGraphAccessTokenAsync();
            var accessToken = await GetUserAccessTokenAsync();
            var response = await SendMessageToTeamsChannelAsync(teamId, channelId, messageContent, accessToken);
            _logger.LogInformation($"graphAPI response: {response}");

            // Step 5: Return the response to the caller
            var responseMessage = req.CreateResponse(response.IsSuccessStatusCode ? System.Net.HttpStatusCode.OK : System.Net.HttpStatusCode.InternalServerError);
            var responseBody = await response.Content.ReadAsStringAsync();

            _logger.LogInformation($"Status Code: {response.StatusCode}, body: {responseBody}");

            await responseMessage.WriteStringAsync(responseBody);

            return responseMessage;
        }

        public async Task<string> GetUserAccessTokenAsync()
        {
            var app = PublicClientApplicationBuilder.Create(CLIENT_ID)
                .WithAuthority(AzureCloudInstance.AzurePublic, TENANT_ID)
                .WithRedirectUri("http://localhost")
                .Build();

            var scopes = new[] { "Group.ReadWrite.All", "ChatMessage.Send" };

            try
            {
                var accounts = await app.GetAccountsAsync();
                var result = await app.AcquireTokenInteractive(scopes)
                    .WithAccount(accounts.FirstOrDefault())
                    .ExecuteAsync();

                return result.AccessToken;
            }
            catch (MsalUiRequiredException)
            {
                // Handle case where user needs to sign in
                var result = await app.AcquireTokenInteractive(scopes)
                    .ExecuteAsync();

                return result.AccessToken;
            }
        }


        // Helper method to get an access token from Microsoft Identity platform
        private async Task<string> GetGraphAccessTokenAsync()
        {
            _logger.LogInformation($"tenantId: {TENANT_ID}");
            _logger.LogInformation($"clientId: {CLIENT_ID}");
            _logger.LogInformation($"clientSecret: {CLIENT_SECRET}");
            var app = ConfidentialClientApplicationBuilder.Create(CLIENT_ID)
                .WithClientSecret(CLIENT_SECRET)
                .WithAuthority(new Uri($"https://login.microsoftonline.com/{TENANT_ID}/"))
                .Build();

            var result = await app.AcquireTokenForClient(new[] { GRAPH_API_SCOPE }).ExecuteAsync();

            return result.AccessToken;
        }

        // Helper method to send a message to a Microsoft Teams channel via Graph API
        private async Task<HttpResponseMessage> SendMessageToTeamsChannelAsync(string teamId, string channelId, string message, string accessToken)
        {
            var apiUrl = GRAPH_API_URL.Replace("{team-id}", teamId).Replace("{channel-id}", channelId);

            var messageBody = new
            {
                body = new
                {
                    content = message
                }
            };

            var jsonContent = new StringContent(JsonSerializer.Serialize(messageBody), System.Text.Encoding.UTF8, "application/json");
            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

            _logger.LogInformation($"Access Token: {accessToken}");

            return await _httpClient.PostAsync(apiUrl, jsonContent);
        }

        // Helper method to look up the teamId and channelId based on the signature
        private (string, string) LookupTeamAndChannelBySignature(string signature)
        {
            // Simulating a lookup table. You can replace this with a database or external lookup.
            var lookupTable = new Dictionary<string, (string teamId, string channelId)>
            {
                { "signature1", ("d2036027-d4d7-47be-a1ad-88047208c5fc", "19:b9a758da4223428cb3a9333e138a2cd5@thread.tacv2") },
                { "signature2", ("team2-id", "channel2-id") },
                { "signature3", ("team3-id", "channel3-id") }
            };

            return lookupTable.ContainsKey(signature) ? lookupTable[signature] : (null, null);
        }
    }
}
