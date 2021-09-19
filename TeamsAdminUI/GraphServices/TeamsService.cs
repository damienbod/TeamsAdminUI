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
            DateTimeOffset end)
        {

            var onlineMeeting = new OnlineMeeting
            {
                StartDateTime = begin,
                EndDateTime = end,
                Subject = meeting,
                LobbyBypassSettings = new LobbyBypassSettings
                {
                    Scope = LobbyBypassScope.Everyone
                }
            };

            return onlineMeeting;
        }

        public OnlineMeeting AddMeetingParticipants(OnlineMeeting onlineMeeting, List<string> attendees)
        {
            var meetingAttendees = new List<MeetingParticipantInfo>();
            foreach (var attendee in attendees)
            {
                if (!string.IsNullOrEmpty(attendee))
                {
                    meetingAttendees.Add(new MeetingParticipantInfo
                    {
                        Upn = attendee.Trim()
                    });
                }
            }

            if (onlineMeeting.Participants == null)
            {
                onlineMeeting.Participants = new MeetingParticipants();
            };
            onlineMeeting.Participants.Attendees = meetingAttendees;

            return onlineMeeting;
        }
    }
}
