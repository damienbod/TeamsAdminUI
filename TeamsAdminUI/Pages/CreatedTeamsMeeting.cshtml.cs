using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.Threading.Tasks;
using TeamsAdminUI.GraphServices;
using Microsoft.Graph;

namespace TeamsAdminUI.Pages
{
    public class CreatedTeamsMeetingModel : PageModel
    {
        private readonly AadGraphApiDelegatedClient _aadGraphApiDelegatedClient;
        private readonly EmailService _emailService;

        public CreatedTeamsMeetingModel(
            AadGraphApiDelegatedClient aadGraphApiDelegatedClient,
            EmailService emailService)
        {
            _aadGraphApiDelegatedClient = aadGraphApiDelegatedClient;
            _emailService = emailService;
        }

        [BindProperty]
        public OnlineMeeting Meeting {get;set;}

        public async Task<ActionResult> OnGetAsync(string meetingId)
        {
            Meeting = await _aadGraphApiDelegatedClient.GetOnlineMeeting(meetingId);
            return Page();
        }

        public async Task<IActionResult> OnPostAsync(string meetingId)
        {
            Meeting = await _aadGraphApiDelegatedClient.GetOnlineMeeting(meetingId);
            foreach (var attendee in Meeting.Participants.Attendees)
            {
                var message = _emailService.CreateStandardEmail(attendee.Upn, Meeting.Subject, Meeting.JoinUrl);
                await _aadGraphApiDelegatedClient.SendEmailAsync(message);
            }
            
            return Page();
        }

    }
}
