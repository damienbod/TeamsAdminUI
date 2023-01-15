using Microsoft.Graph;

namespace TeamsAdminUIObo.GraphServices;

public class AadGraphApiApplicationClient
{
    private readonly IConfiguration _configuration;
    private readonly GraphApplicationClientService _graphApplicationClientService;

    public AadGraphApiApplicationClient(IConfiguration configuration, 
        GraphApplicationClientService graphApplicationClientService)
    {
        _configuration = configuration;
        _graphApplicationClientService = graphApplicationClientService;
    }

    private async Task<string> GetUserIdAsync()
    {
        var meetingOrganizer = _configuration["AzureAd:MeetingOrganizer"];
        var filter = $"startswith(userPrincipalName,'{meetingOrganizer}')";
        var graphServiceClient = _graphApplicationClientService.GetGraphClientWithManagedIdentityOrDevClient();

        var users = await graphServiceClient.Users
            .Request()
            .Filter(filter)
            .GetAsync();

        return users.CurrentPage[0].Id;
    }

    public async Task SendEmailAsync(Message message)
    {
        var graphServiceClient = _graphApplicationClientService.GetGraphClientWithManagedIdentityOrDevClient();

        var saveToSentItems = true;

        var userId = await GetUserIdAsync();

        await graphServiceClient.Users[userId]
            .SendMail(message, saveToSentItems)
            .Request()
            .PostAsync();
    }

    public async Task<OnlineMeeting> CreateOnlineMeeting(OnlineMeeting onlineMeeting)
    {
        var graphServiceClient = _graphApplicationClientService.GetGraphClientWithManagedIdentityOrDevClient();

        var userId = await GetUserIdAsync();

        return await graphServiceClient.Users[userId]
            .OnlineMeetings
            .Request()
            .AddAsync(onlineMeeting);
    }

    public async Task<OnlineMeeting> UpdateOnlineMeeting(OnlineMeeting onlineMeeting)
    {
        var graphServiceClient = _graphApplicationClientService.GetGraphClientWithManagedIdentityOrDevClient();

        var userId = await GetUserIdAsync();

        return await graphServiceClient.Users[userId]
            .OnlineMeetings[onlineMeeting.Id]
            .Request()
            .UpdateAsync(onlineMeeting);
    }

    public async Task<OnlineMeeting> GetOnlineMeeting(string onlineMeetingId)
    {
        var graphServiceClient = _graphApplicationClientService.GetGraphClientWithManagedIdentityOrDevClient();

        var userId = await GetUserIdAsync();

        return await graphServiceClient.Users[userId]
            .OnlineMeetings[onlineMeetingId]
            .Request()
            .GetAsync();
    }
}
