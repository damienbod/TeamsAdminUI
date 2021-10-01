using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using System.Threading.Tasks;

namespace TeamsAdminUIObo.GraphServices
{
    public class AadGraphApiApplicationClient
    {
        private readonly ApiTokenInMemoryClient _apiTokenInMemoryClient;
        private readonly IConfiguration _configuration;

        public AadGraphApiApplicationClient(ApiTokenInMemoryClient apiTokenInMemoryClient,
            IConfiguration configuration)
        {
            _apiTokenInMemoryClient = apiTokenInMemoryClient;
            _configuration = configuration;
        }

        private async Task<string> GetUserIdAsync()
        {
            var meetingOrganizer = _configuration["AzureAd:MeetingOrganizer"];
            var filter = $"startswith(userPrincipalName,'{meetingOrganizer}')";
            var graphServiceClient = await _apiTokenInMemoryClient.GetGraphClient();

            var users = await graphServiceClient.Users
                .Request()
                .Filter(filter)
                .GetAsync();

            return users.CurrentPage[0].Id;
        }

        public async Task SendEmailAsync(Message message)
        {
            var graphServiceClient = await _apiTokenInMemoryClient.GetGraphClient();

            var saveToSentItems = true;

            var userId = await GetUserIdAsync();

            await graphServiceClient.Users[userId]
                .SendMail(message, saveToSentItems)
                .Request()
                .PostAsync();
        }

        public async Task<OnlineMeeting> CreateOnlineMeeting(OnlineMeeting onlineMeeting)
        {
            var graphServiceClient = await _apiTokenInMemoryClient.GetGraphClient();

            var userId = await GetUserIdAsync();

            return await graphServiceClient.Users[userId]
                .OnlineMeetings
                .Request()
                .AddAsync(onlineMeeting);
        }

        public async Task<OnlineMeeting> UpdateOnlineMeeting(OnlineMeeting onlineMeeting)
        {
            var graphServiceClient = await _apiTokenInMemoryClient.GetGraphClient();

            var userId = await GetUserIdAsync();

            return await graphServiceClient.Users[userId]
                .OnlineMeetings[onlineMeeting.Id]
                .Request()
                .UpdateAsync(onlineMeeting);
        }

        public async Task<OnlineMeeting> GetOnlineMeeting(string onlineMeetingId)
        {
            var graphServiceClient = await _apiTokenInMemoryClient.GetGraphClient();

            var userId = await GetUserIdAsync();

            return await graphServiceClient.Users[userId]
                .OnlineMeetings[onlineMeetingId]
                .Request()
                .GetAsync();
        }
    }
}
