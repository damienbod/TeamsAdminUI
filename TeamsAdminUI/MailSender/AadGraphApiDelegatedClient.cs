using Microsoft.Graph;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace TeamsAdminUI.MailSender
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

    }
}
