using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Graph;

namespace TeamsAdminUI.Pages
{
    public class CreatedTeamsMeetingModel : PageModel
    {
        private readonly AadGraphApiDelegatedClient _aadGraphApiDelegatedClient;
        private readonly TeamsService _teamsService;

        public CreatedTeamsMeetingModel(
            AadGraphApiDelegatedClient aadGraphApiDelegatedClient,
            TeamsService teamsService)
        {
            _aadGraphApiDelegatedClient = aadGraphApiDelegatedClient;
            _teamsService = teamsService;
        }

        [BindProperty]
        public OnlineMeeting Meeting {get;set;}

        public async Task<ActionResult> OnGetAsync(string meetingId)
        {
            OnlineMeeting = await _aadGraphApiDelegatedClient.GetOnlineMeeting(meetingId);
        }
    }
}
