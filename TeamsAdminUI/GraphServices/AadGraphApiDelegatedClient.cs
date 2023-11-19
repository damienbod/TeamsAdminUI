using Microsoft.Graph;
using Microsoft.Graph.Me.SendMail;
using Microsoft.Graph.Models;

namespace TeamsAdminUI.GraphServices;

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

        var body = new SendMailPostRequestBody
        {
            Message = message,
            SaveToSentItems = saveToSentItems
        };

        await _graphServiceClient.Me.SendMail
            .PostAsync(body);
    }

    public async Task<OnlineMeeting?> CreateOnlineMeeting(OnlineMeeting onlineMeeting)
    {
        return await _graphServiceClient.Me
            .OnlineMeetings.PostAsync(onlineMeeting);
    }

    public async Task<OnlineMeeting?> UpdateOnlineMeeting(OnlineMeeting onlineMeeting)
    {
        return await _graphServiceClient.Me
            .OnlineMeetings[onlineMeeting.Id]
            .PatchAsync(onlineMeeting);
    }

    public async Task<OnlineMeeting?> GetOnlineMeeting(string onlineMeetingId)
    {
        return await _graphServiceClient.Me
            .OnlineMeetings[onlineMeetingId]
            .GetAsync();
    }
}
