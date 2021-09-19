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
        public OnlineMeeting Meeting { get; set; }

        [BindProperty]
        public string EmailSent { get; set; }

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
                var recipient = attendee.Upn.Trim();
                var message = _emailService.CreateStandardEmail(recipient, Meeting.Subject, Meeting.JoinUrl);
                await _aadGraphApiDelegatedClient.SendEmailAsync(message);
            }

            EmailSent = "Emails sent to all attendees, please check your mailbox";
            return Page();
        }

    }
}
