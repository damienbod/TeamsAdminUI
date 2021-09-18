using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Identity.Web;
using System;
using System.Threading.Tasks;
using TeamsAdminUI.GraphServices;

namespace TeamsAdminUI.Pages
{
    [AuthorizeForScopes(Scopes = new string[] { "User.read", "Mail.Send", "Mail.ReadWrite", "OnlineMeetings.ReadWrite" })]
    public class CreateTeamsMeetingModel : PageModel
    {
        private readonly AadGraphApiDelegatedClient _aadGraphApiDelegatedClient;
        private readonly TeamsService _teamsService;

        public string JoinUrl { get; set; }

        public DateTimeOffset Begin { get; set; }
        public DateTimeOffset End { get; set; }
        public string AttendeeEmail { get; set; }
        public string MeetingName { get; set; }

        public CreateTeamsMeetingModel(AadGraphApiDelegatedClient aadGraphApiDelegatedClient,
            TeamsService teamsService)
        {
            _aadGraphApiDelegatedClient = aadGraphApiDelegatedClient;
            _teamsService = teamsService;
        }

        public async Task OnPostAsync()
        {
            var user = new Microsoft.Graph.User
            {
                Id = AttendeeEmail
            };
            var meeting = _teamsService.CreateTeamsMeeting(MeetingName, Begin, End, user);
            var createdMeeting = await _aadGraphApiDelegatedClient.CreateOnlineMeeting(meeting);

            JoinUrl = createdMeeting.JoinUrl;
        }

        public void OnGet()
        {
            Begin = DateTimeOffset.UtcNow;
            End = DateTimeOffset.UtcNow.AddMinutes(60);
        }
    }
}
