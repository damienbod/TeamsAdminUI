using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Identity.Web;
using System;
using System.Collections.Generic;
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

        [BindProperty]
        public DateTimeOffset Begin { get; set; }
        [BindProperty]
        public DateTimeOffset End { get; set; }
        [BindProperty]
        public string AttendeeEmail { get; set; }
        [BindProperty]
        public string MeetingName { get; set; }

        public CreateTeamsMeetingModel(AadGraphApiDelegatedClient aadGraphApiDelegatedClient,
            TeamsService teamsService)
        {
            _aadGraphApiDelegatedClient = aadGraphApiDelegatedClient;
            _teamsService = teamsService;
        }

        public async Task<IActionResult> OnPostAsync()
        {
            if (!ModelState.IsValid)
            {
                return Page();
            }

            var meeting = _teamsService.CreateTeamsMeeting(MeetingName, Begin, End);

            var updatedMeeting = _teamsService.AddMeetingParticipants(
              meeting, new List<string> { AttendeeEmail });

            var createdMeeting = await _aadGraphApiDelegatedClient.CreateOnlineMeeting(updatedMeeting);

            JoinUrl = createdMeeting.JoinUrl;

          

            //var result = await _aadGraphApiDelegatedClient
            //    .UpdateOnlineMeeting(meetingToUpdate);

            return RedirectToPage("./CreatedTeamsMeeting", "Get", new { meetingId = createdMeeting.Id });
        }

        public void OnGet()
        {
            Begin = DateTimeOffset.UtcNow;
            End = DateTimeOffset.UtcNow.AddMinutes(60);
        }
    }
}
