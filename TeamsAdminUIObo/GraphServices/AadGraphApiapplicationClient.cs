using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using System.Threading.Tasks;

namespace TeamsAdminUIObo.GraphServices
{
	public class AadGraphApiapplicationClient
    {
		private readonly ApiTokenInMemoryClient _apiTokenInMemoryClient;
		private readonly ILogger<AadGraphApiapplicationClient> _logger;

		public AadGraphApiapplicationClient(ApiTokenInMemoryClient apiTokenInMemoryClient,
            ILoggerFactory loggerFactory)
        {
            _apiTokenInMemoryClient = apiTokenInMemoryClient;
            _logger = loggerFactory.CreateLogger<AadGraphApiapplicationClient>();
        }

        private async Task<string> GetUserIdAsync()
		{
            var meetingOrganizer = "damienbod@damienbodsharepoint.onmicrosoft.com";
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

        void MyLoggingMethod(Microsoft.Identity.Client.LogLevel level, string message, bool containsPii)
        {
            _logger.LogInformation($"MSAL {level} {containsPii} {message}");
        }
    }

}
