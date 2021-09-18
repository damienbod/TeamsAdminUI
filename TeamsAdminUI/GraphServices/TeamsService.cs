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
                //Participants = new MeetingParticipants
                //{
                //    Attendees = new List<MeetingParticipantInfo>()
                //    {
                //        new MeetingParticipantInfo
                //        {
                //            Identity = new IdentitySet
                //            {
                //                User = new Identity
                //                {
                //                    Id = attendee.Id
                //                }
                //            },
                //            Upn = attendee.UserPrincipalName
                //        }
                //    }
                //}
            };

            return onlineMeeting;
        }

        public OnlineMeeting AddMeetingParticipants(OnlineMeeting onlineMeeting, List<string> attendees)
        {
            var meetingAttendees = new List<MeetingParticipantInfo>();
            foreach(var attendee in attendees)
            {
                meetingAttendees.Add(new MeetingParticipantInfo
                {
                    Upn = attendee
                });
            }

            if(onlineMeeting.Participants == null)
            {
                onlineMeeting.Participants = new MeetingParticipants();
            };
            onlineMeeting.Participants.Attendees = meetingAttendees;

            return onlineMeeting;
        }
    }
}
