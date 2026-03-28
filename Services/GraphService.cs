using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace TeamsCallApi.Services;

public class GraphService
{
    private readonly GraphServiceClient _graphClient;
    private readonly IConfiguration _configuration;
    private readonly ILogger<GraphService> _logger;

    public GraphService(IConfiguration configuration, ILogger<GraphService> logger)
    {
        _configuration = configuration;
        _logger = logger;

        var tenantId     = _configuration["AzureAd:TenantId"]!;
        var clientId     = _configuration["AzureAd:ClientId"]!;
        var clientSecret = _configuration["AzureAd:ClientSecret"]!;

        var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);

        _graphClient = new GraphServiceClient(credential);
    }

    public async Task<Call?> InitiateCallAsync(string targetUserId, string callbackUrl, string botId)
    {
        _logger.LogInformation("[GraphService] Initiating call to user {UserId}", targetUserId);

        var tenantId = _configuration["AzureAd:TenantId"]!;

        var call = new Call
        {
            Direction = CallDirection.Outgoing,
            CallbackUri = callbackUrl,
            RequestedModalities = new List<Modality?> { Modality.Audio },
            TenantId = tenantId,

            // ✅ CRITICAL: Required for Graph calling
            MediaConfig = new ServiceHostedMediaConfig(),

            // ✅ BOT (application identity)
            Source = new ParticipantInfo
            {
                Identity = new IdentitySet
                {
                    Application = new Identity
                    {
                        Id = botId,
                        DisplayName = "TeamsCallBot"
                    }
                }
            },

            // ✅ TARGET USER
            Targets = new List<InvitationParticipantInfo>
            {
                new InvitationParticipantInfo
                {
                    Identity = new IdentitySet
                    {
                        User = new Identity
                        {
                            Id = targetUserId
                        }
                    }
                }
            }
        };

        try
        {
            var createdCall = await _graphClient.Communications.Calls.PostAsync(call);

            _logger.LogInformation("[GraphService] Call created successfully. CallId: {CallId}", createdCall?.Id);

            return createdCall;
        }
        catch (Microsoft.Graph.Models.ODataErrors.ODataError ex)
        {
            _logger.LogError("[GraphService] Graph API error: {Code} - {Message}",
                ex.Error?.Code, ex.Error?.Message);

            throw;
        }
        catch (Exception ex)
        {
            _logger.LogError("[GraphService] Unexpected error: {Message}", ex.Message);
            throw;
        }
    }
}