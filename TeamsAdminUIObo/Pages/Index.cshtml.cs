using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Identity.Web;
using TeamsAdminUIObo.GraphServices;

namespace TeamsAdminUIObo.Pages;

[AuthorizeForScopes(Scopes = new string[] { "User.read", "Mail.Send", "Mail.ReadWrite", "OnlineMeetings.ReadWrite" })]
public class CreateTeamsMeetingModel : PageModel
{
    private readonly AadGraphApiApplicationClient _aadGraphApiDelegatedClient;
    private readonly TeamsService _teamsService;

    public string? JoinUrl { get; set; }

    [BindProperty]
    public DateTimeOffset Begin { get; set; } = DateTimeOffset.Now;
    [BindProperty]
    public DateTimeOffset End { get; set; } = DateTimeOffset.Now;
    [BindProperty]
    public string? AttendeeEmail { get; set; }
    [BindProperty]
    public string MeetingName { get; set; } = string.Empty;

    public CreateTeamsMeetingModel(AadGraphApiApplicationClient aadGraphApiDelegatedClient,
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

        var attendees = AttendeeEmail!.Split(';');
        List<string> items = new();
        items.AddRange(attendees);
        var updatedMeeting = _teamsService.AddMeetingParticipants(
          meeting, items);

        var createdMeeting = await _aadGraphApiDelegatedClient.CreateOnlineMeeting(updatedMeeting);

        JoinUrl = createdMeeting.JoinWebUrl;

        return RedirectToPage("./CreatedTeamsMeeting", "Get", new { meetingId = createdMeeting.Id });
    }

    public void OnGet()
    {
        Begin = DateTimeOffset.UtcNow;
        End = DateTimeOffset.UtcNow.AddMinutes(60);
    }
}
