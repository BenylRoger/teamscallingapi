using Microsoft.AspNetCore.Mvc;
using TeamsCallApi.Services;

namespace TeamsCallApi.Controllers;

[ApiController]
[Route("api/call")]
public class CallController : ControllerBase
{
    private readonly ILogger<CallController> _logger;
    private readonly IConfiguration _configuration;
    private readonly GraphService _graphService;

    public CallController(
        ILogger<CallController> logger,
        IConfiguration configuration,
        GraphService graphService)
    {
        _logger = logger;
        _configuration = configuration;
        _graphService = graphService;
    }

    /// <summary>
    /// POST /api/call/make-call
    /// Initiates a real Microsoft Teams call via Microsoft Graph API.
    /// </summary>
    [HttpPost("make-call")]
    public async Task<IActionResult> MakeCall()
    {
        var callbackUrl  = _configuration["Graph:CallbackUrl"]!;
        var botId        = _configuration["Graph:BotId"]!;
        var targetUserId = _configuration["Graph:TargetUserId"]!;

        _logger.LogInformation("[MakeCall] Initiating Teams call to user {UserId}", targetUserId);

        try
        {
            var call = await _graphService.InitiateCallAsync(targetUserId, callbackUrl, botId);

            return Ok(new
            {
                Status      = "Call initiated successfully",
                CallId      = call?.Id,
                State       = call?.State?.ToString(),
                InitiatedAt = DateTime.UtcNow
            });
        }
        catch (Microsoft.Graph.Models.ODataErrors.ODataError ex)
        {
            _logger.LogError("[MakeCall] Graph API error: {Code} - {Message}",
                ex.Error?.Code, ex.Error?.Message);

            return StatusCode(500, new
            {
                Error   = "Graph API call failed",
                Code    = ex.Error?.Code,
                Message = ex.Error?.Message
            });
        }
        catch (Exception ex)
        {
            _logger.LogError("[MakeCall] Unexpected error: {Message}", ex.Message);
            return StatusCode(500, new { Error = ex.Message });
        }
    }
}
