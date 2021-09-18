using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Identity.Web;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using TeamsAdminUI.GraphServices;
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
            Meeting = await _aadGraphApiDelegatedClient.GetOnlineMeeting(meetingId);
            return Page();
        }
    }
}
