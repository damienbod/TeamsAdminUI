using Microsoft.Graph;
using System;
using System.Collections.Generic;

namespace TeamsAdminUI.GraphServices
{
    public class TeamsService
    {
        public OnlineMeeting CreateTeamsMeeting(
            string meeting, 
            DateTimeOffset begin, 
            DateTimeOffset end,
            User attendee)
        {

            var onlineMeeting = new OnlineMeeting
            {
                StartDateTime = begin,
                EndDateTime = end,
                Subject = meeting,
                Participants = new MeetingParticipants
                {
                    Attendees = new List<MeetingParticipantInfo>()
                    {
                        new MeetingParticipantInfo
                        {
                            Identity = new IdentitySet
                            {
                                User = new Identity
                                {
                                    Id = attendee.Id
                                }
                            },
                            Upn = attendee.UserPrincipalName
                        }
                    }
                }
            };

            return onlineMeeting;
        }

    }
}
