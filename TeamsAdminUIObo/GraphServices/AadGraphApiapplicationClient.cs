﻿using Microsoft.Graph;
using Microsoft.Graph.Me.SendMail;
using Microsoft.Graph.Models;

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

        var userId = await GetUserIdAsync();
        var saveToSentItems = true;

        var body = new SendMailPostRequestBody
        {
            Message = message,
            SaveToSentItems = saveToSentItems
        };

        await graphServiceClient.Me.SendMail
            .PostAsync(body);
    }

    public async Task<OnlineMeeting?> CreateOnlineMeeting(OnlineMeeting onlineMeeting)
    {
        var graphServiceClient = _graphApplicationClientService.GetGraphClientWithManagedIdentityOrDevClient();

        var userId = await GetUserIdAsync();

        return await graphServiceClient.Users[userId]
            .OnlineMeetings
            .PostAsync(onlineMeeting);
    }

    public async Task<OnlineMeeting?> UpdateOnlineMeeting(OnlineMeeting onlineMeeting)
    {
        var graphServiceClient = _graphApplicationClientService.GetGraphClientWithManagedIdentityOrDevClient();

        var userId = await GetUserIdAsync();

        return await graphServiceClient.Users[userId]
            .OnlineMeetings[onlineMeeting.Id]
            .PatchAsync(onlineMeeting);
    }

    public async Task<OnlineMeeting?> GetOnlineMeeting(string onlineMeetingId)
    {
        var graphServiceClient = _graphApplicationClientService.GetGraphClientWithManagedIdentityOrDevClient();

        var userId = await GetUserIdAsync();

        return await graphServiceClient.Users[userId]
            .OnlineMeetings[onlineMeetingId]
            .GetAsync();
    }
}
