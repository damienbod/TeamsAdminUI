using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc.Authorization;
using Microsoft.Identity.Web;
using Microsoft.Identity.Web.UI;
using Microsoft.IdentityModel.Logging;
using TeamsAdminUI.GraphServices;

namespace TeamsAdminUI;

public class Program
{
    public static void Main(string[] args)
    {
        var builder = WebApplication.CreateBuilder(args);

        builder.Services.AddScoped<AadGraphApiDelegatedClient>();
        builder.Services.AddScoped<EmailService>();
        builder.Services.AddScoped<TeamsService>();
        builder.Services.AddHttpClient();
        builder.Services.AddOptions();

        var scopes = new List<string>()
        {
            "User.read",
            "Mail.Send",
            "Mail.ReadWrite",
            "OnlineMeetings.ReadWrite"
        };

        builder.Services.AddMicrosoftIdentityWebAppAuthentication(builder.Configuration)
            .EnableTokenAcquisitionToCallDownstreamApi()
            .AddMicrosoftGraph("https://graph.microsoft.com/v1.0", scopes)
            .AddInMemoryTokenCaches();

        builder.Services.AddRazorPages().AddMvcOptions(options =>
        {
            var policy = new AuthorizationPolicyBuilder()
                .RequireAuthenticatedUser()
                .Build();
            options.Filters.Add(new AuthorizeFilter(policy));
        }).AddMicrosoftIdentityUI();

        var app = builder.Build();

        IdentityModelEventSource.ShowPII = true;

        if (!app.Environment.IsDevelopment())
        {
            app.UseExceptionHandler("/Error");
            app.UseHsts();
        }

        app.UseHttpsRedirection();
        app.UseStaticFiles();

        app.UseRouting();

        app.UseAuthentication();
        app.UseAuthorization();

        app.MapRazorPages();
        app.MapControllers();

        app.Run();
    }
}