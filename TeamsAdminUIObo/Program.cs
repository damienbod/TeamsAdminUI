using TeamsAdminUIObo.GraphServices;

namespace TeamsAdminUIObo;

public class Program
{
    public static void Main(string[] args)
    {
        var builder = WebApplication.CreateBuilder(args);

        builder.Services.AddSingleton<GraphApplicationClientService>();
        builder.Services.AddScoped<AadGraphApiApplicationClient>();
        builder.Services.AddScoped<EmailService>();
        builder.Services.AddScoped<TeamsService>();
        builder.Services.AddHttpClient();
        builder.Services.AddOptions();

        builder.Services.AddRazorPages();

        var app = builder.Build();

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