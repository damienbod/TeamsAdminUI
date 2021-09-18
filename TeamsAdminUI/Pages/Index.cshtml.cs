using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Identity.Web;
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

        public void  OnGet()
        {
        }
    }
}