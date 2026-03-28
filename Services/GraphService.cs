using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Communications.Calls.Item.Answer;
using Microsoft.Graph.Communications.Calls.Item.Reject;

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

    // ─── Initiate outgoing call ───────────────────────────────────────────────

    public async Task<Call?> InitiateCallAsync(string targetUserId, string callbackUrl, string botId)
    {
        _logger.LogInformation("[GraphService] Initiating call to user {UserId}", targetUserId);

        var tenantId = _configuration["AzureAd:TenantId"]!;

        var call = new Call
        {
            Direction            = CallDirection.Outgoing,
            CallbackUri          = callbackUrl,
            RequestedModalities  = new List<Modality?> { Modality.Audio },
            TenantId             = tenantId,
            MediaConfig          = new ServiceHostedMediaConfig(),

            Source = new ParticipantInfo
            {
                Identity = new IdentitySet
                {
                    Application = new Identity
                    {
                        Id          = botId,
                        DisplayName = "TeamsCallBot"
                    }
                }
            },

            Targets = new List<InvitationParticipantInfo>
            {
                new InvitationParticipantInfo
                {
                    Identity = new IdentitySet
                    {
                        User = new Identity { Id = targetUserId }
                    }
                }
            }
        };

        try
        {
            var createdCall = await _graphClient.Communications.Calls.PostAsync(call);
            _logger.LogInformation("[GraphService] Call created. CallId: {CallId}", createdCall?.Id);
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

    // ─── Answer an incoming call ──────────────────────────────────────────────

    public async Task AnswerCallAsync(string callId)
    {
        _logger.LogInformation("[GraphService] Answering call {CallId}", callId);

        try
        {
            var requestBody = new AnswerPostRequestBody
            {
                CallbackUri = _configuration["Graph:CallbackUrl"]!,
                MediaConfig = new ServiceHostedMediaConfig(),
                AcceptedModalities = new List<Modality?> { Modality.Audio }
            };

            await _graphClient.Communications.Calls[callId].Answer.PostAsync(requestBody);
            _logger.LogInformation("[GraphService] Call {CallId} answered successfully", callId);
        }
        catch (Microsoft.Graph.Models.ODataErrors.ODataError ex)
        {
            _logger.LogError("[GraphService] Answer error: {Code} - {Message}",
                ex.Error?.Code, ex.Error?.Message);
            throw;
        }
    }

    // ─── Reject an incoming call ──────────────────────────────────────────────

    public async Task RejectCallAsync(string callId)
    {
        _logger.LogInformation("[GraphService] Rejecting call {CallId}", callId);

        try
        {
            var requestBody = new RejectPostRequestBody
            {
                Reason = RejectReason.Busy
            };

            await _graphClient.Communications.Calls[callId].Reject.PostAsync(requestBody);
            _logger.LogInformation("[GraphService] Call {CallId} rejected", callId);
        }
        catch (Microsoft.Graph.Models.ODataErrors.ODataError ex)
        {
            _logger.LogError("[GraphService] Reject error: {Code} - {Message}",
                ex.Error?.Code, ex.Error?.Message);
            throw;
        }
    }

    // ─── Hang up / terminate a call ──────────────────────────────────────────

    public async Task HangUpCallAsync(string callId)
    {
        _logger.LogInformation("[GraphService] Hanging up call {CallId}", callId);

        try
        {
            await _graphClient.Communications.Calls[callId].DeleteAsync();
            _logger.LogInformation("[GraphService] Call {CallId} terminated", callId);
        }
        catch (Microsoft.Graph.Models.ODataErrors.ODataError ex)
        {
            _logger.LogError("[GraphService] HangUp error: {Code} - {Message}",
                ex.Error?.Code, ex.Error?.Message);
            throw;
        }
    }
}