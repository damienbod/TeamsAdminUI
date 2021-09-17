using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Identity.Web;
using System;
using System.Threading.Tasks;
using TeamsAdminUI.GraphServices;

namespace TeamsAdminUI.Pages
{
    [AuthorizeForScopes(Scopes = new string[] { "User.read", "Mail.Send", "Mail.ReadWrite", "OnlineMeetings.ReadWrite" })]
    public class CallApiModel : PageModel
    {
        private readonly AadGraphApiDelegatedClient _aadGraphApiDelegatedClient;
        private readonly TeamsService _teamsService;

        public string JoinUrl { get; set; }
        public CallApiModel(AadGraphApiDelegatedClient aadGraphApiDelegatedClient,
            TeamsService teamsService)
        {
            _aadGraphApiDelegatedClient = aadGraphApiDelegatedClient;
            _teamsService = teamsService;
        }

        public async Task OnGetAsync()
        {
            var begin = DateTimeOffset.UtcNow;
            var end = DateTimeOffset.UtcNow.AddMinutes(60);
            var meeting = _teamsService.CreateTeamsMeeting("my meeting", begin, end);
            var createdMeeting = await _aadGraphApiDelegatedClient.CreateOnlineMeeting(meeting);

            JoinUrl = createdMeeting.JoinUrl;
        }
    }
}