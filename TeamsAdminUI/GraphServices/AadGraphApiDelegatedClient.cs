using Microsoft.Graph;
using System.Threading.Tasks;

namespace TeamsAdminUI.GraphServices
{
    public class AadGraphApiDelegatedClient
    {
        private readonly GraphServiceClient _graphServiceClient;

        public AadGraphApiDelegatedClient(GraphServiceClient graphServiceClient)
        {
            _graphServiceClient = graphServiceClient;
        }

        public async Task SendEmailAsync(Message message)
        {
            var saveToSentItems = true;

            await _graphServiceClient.Me
                .SendMail(message, saveToSentItems)
                .Request()
                .PostAsync();
        }

        public async Task<OnlineMeeting> CreateOnlineMeeting(OnlineMeeting onlineMeeting)
        {
            return await _graphServiceClient.Me
                .OnlineMeetings
                .Request()
                .AddAsync(onlineMeeting);
        }

        public async Task<OnlineMeeting> UpdateOnlineMeeting(OnlineMeeting onlineMeeting)
        {
            return await _graphServiceClient.Me
                .OnlineMeetings[onlineMeeting.Id]
                .Request()
                .UpdateAsync(onlineMeeting);
        }

        public async Task<OnlineMeeting> GetOnlineMeeting(string onlineMeetingId)
        {
            return await _graphServiceClient.Me
                .OnlineMeetings[onlineMeetingId]
                .Request()
                .GetAsync();
        }
    }



}
