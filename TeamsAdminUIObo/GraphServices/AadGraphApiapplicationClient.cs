using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using System.Threading.Tasks;

namespace TeamsAdminUIObo.GraphServices
{
    public class AadGraphApiApplicationClient
    {
        private readonly IConfiguration _configuration;

        public AadGraphApiApplicationClient(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        private async Task<string> GetUserIdAsync()
        {
            var meetingOrganizer = _configuration["AzureAd:MeetingOrganizer"];
            var filter = $"startswith(userPrincipalName,'{meetingOrganizer}')";
            var graphServiceClient = await GetGraphClient();

            var users = await graphServiceClient.Users
                .Request()
                .Filter(filter)
                .GetAsync();

            return users.CurrentPage[0].Id;
        }

        public async Task SendEmailAsync(Message message)
        {
            var graphServiceClient = await GetGraphClient();

            var saveToSentItems = true;

            var userId = await GetUserIdAsync();

            await graphServiceClient.Users[userId]
                .SendMail(message, saveToSentItems)
                .Request()
                .PostAsync();
        }

        public async Task<OnlineMeeting> CreateOnlineMeeting(OnlineMeeting onlineMeeting)
        {
            var graphServiceClient = await GetGraphClient();

            var userId = await GetUserIdAsync();

            return await graphServiceClient.Users[userId]
                .OnlineMeetings
                .Request()
                .AddAsync(onlineMeeting);
        }

        public async Task<OnlineMeeting> UpdateOnlineMeeting(OnlineMeeting onlineMeeting)
        {
            var graphServiceClient = await GetGraphClient();

            var userId = await GetUserIdAsync();

            return await graphServiceClient.Users[userId]
                .OnlineMeetings[onlineMeeting.Id]
                .Request()
                .UpdateAsync(onlineMeeting);
        }

        public async Task<OnlineMeeting> GetOnlineMeeting(string onlineMeetingId)
        {
            var graphServiceClient = await GetGraphClient();

            var userId = await GetUserIdAsync();

            return await graphServiceClient.Users[userId]
                .OnlineMeetings[onlineMeetingId]
                .Request()
                .GetAsync();
        }

        private async Task<GraphServiceClient> GetGraphClient()
        {

            string[] scopes = new[] { "https://graph.microsoft.com/.default" };
            var tenantId = _configuration["AzureAd:TenantId"];

            // Values from app registration
            var clientId = _configuration.GetValue<string>("AzureAd:ClientId");
            var clientSecret = _configuration.GetValue<string>("AzureAd:ClientSecret");

            var options = new TokenCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
            };

            // https://docs.microsoft.com/dotnet/api/azure.identity.clientsecretcredential
            var clientSecretCredential = new ClientSecretCredential(
                tenantId, clientId, clientSecret, options);

            return new GraphServiceClient(clientSecretCredential, scopes);
        }
    }
}
