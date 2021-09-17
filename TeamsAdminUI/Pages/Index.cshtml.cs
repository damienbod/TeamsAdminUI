using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Identity.Web;
using System.Threading.Tasks;
using TeamsAdminUI.MailSender;

namespace TeamsAdminUI.Pages
{
    public class CallApiModel : PageModel
    {
        private readonly AadGraphApiDelegatedClient _aadGraphApiDelegatedClient;

        public CallApiModel(AadGraphApiDelegatedClient aadGraphApiDelegatedClient)
        {
            _aadGraphApiDelegatedClient = aadGraphApiDelegatedClient;
        }

        public void OnGet()
        {
           
        }
    }
}