using Azure.Identity;
using Microsoft.Graph;

namespace TeamsAdminUIObo.GraphServices;

public class GraphApplicationClientService
{
    private readonly IConfiguration _configuration;
    private readonly IHostEnvironment _environment;
    private GraphServiceClient? _graphServiceClient;

    public GraphApplicationClientService(IConfiguration configuration, IHostEnvironment environment)
    {
        _configuration = configuration;
        _environment = environment;
    }

    /// <summary>
    /// gets a singleton instance of the GraphServiceClient
    /// </summary>
    /// <returns></returns>
    public GraphServiceClient GetGraphClientWithManagedIdentityOrDevClient()
    {
        if (_graphServiceClient != null)
            return _graphServiceClient;

        string[] scopes = new[] { "https://graph.microsoft.com/.default" };

        var chainedTokenCredential = GetChainedTokenCredentials();
        _graphServiceClient = new GraphServiceClient(chainedTokenCredential, scopes);

        return _graphServiceClient;
    }

    private ChainedTokenCredential GetChainedTokenCredentials()
    {
        if (!_environment.IsDevelopment())
        {
            return new ChainedTokenCredential(new ManagedIdentityCredential());
        }
        else // dev env
        {
            var tenantId = _configuration["AzureAd:TenantId"];
            var clientId = _configuration.GetValue<string>("AzureAd:ClientId");
            var clientSecret = _configuration.GetValue<string>("AzureAd:ClientSecret");

            var options = new TokenCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
            };

            // https://docs.microsoft.com/dotnet/api/azure.identity.clientsecretcredential
            var devClientSecretCredential = new ClientSecretCredential(
                tenantId, clientId, clientSecret, options);

            var chainedTokenCredential = new ChainedTokenCredential(devClientSecretCredential);

            return chainedTokenCredential;
        }
    }
}