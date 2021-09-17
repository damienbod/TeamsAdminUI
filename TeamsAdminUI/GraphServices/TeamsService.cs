using Microsoft.Graph;
using System;

namespace TeamsAdminUI.GraphServices
{
    public class TeamsService
    {
        public OnlineMeeting CreateTeamsMeeting(
            string meeting, 
            DateTimeOffset begin, 
            DateTimeOffset end)
        {
            var onlineMeeting = new OnlineMeeting
            {
                StartDateTime = begin,
                EndDateTime = end,
                Subject = meeting
            };

            return onlineMeeting;
        }

    }
}
