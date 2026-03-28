using Microsoft.AspNetCore.Mvc;
using Microsoft.Graph.Communications.Calls.Item.Reject;
using TeamsCallApi.Models;
using TeamsCallApi.Services;

namespace TeamsCallApi.Controllers;

[ApiController]
[Route("api/calls")]
public class CallbackController : ControllerBase
{
    private readonly ILogger<CallbackController> _logger;
    private readonly GraphService _graphService;
    private readonly IConfiguration _configuration;

    public CallbackController(
        ILogger<CallbackController> logger,
        GraphService graphService,
        IConfiguration configuration)
    {
        _logger        = logger;
        _graphService  = graphService;
        _configuration = configuration;
    }

    /// <summary>
    /// POST /api/calls
    /// Receives all call state notifications from Microsoft Graph.
    /// Handles: incoming, establishing, established, terminated
    /// </summary>
    [HttpPost]
    public async Task<IActionResult> HandleCallback([FromBody] CallbackPayload payload)
    {
        if (payload?.Value == null || !payload.Value.Any())
        {
            _logger.LogWarning("[Callback] Empty or null payload received");
            return Ok();
        }

        foreach (var notification in payload.Value)
        {
            var callId    = notification.ResourceData?.Id;
            var state     = notification.ResourceData?.State?.ToLower();
            var direction = notification.ResourceData?.Direction?.ToLower();

            _logger.LogInformation("=== [Graph Callback] ===");
            _logger.LogInformation("Change Type : {ChangeType}", notification.ChangeType);
            _logger.LogInformation("Call ID     : {CallId}", callId);
            _logger.LogInformation("State       : {State}", state);
            _logger.LogInformation("Direction   : {Direction}", direction);
            _logger.LogInformation("========================");

            if (string.IsNullOrEmpty(callId)) continue;

            // Handle each call state
            switch (state)
            {
                case "incoming":
                    await HandleIncomingCall(callId, direction);
                    break;

                case "establishing":
                    _logger.LogInformation("[Callback] Call {CallId} is establishing...", callId);
                    break;

                case "established":
                    _logger.LogInformation("[Callback] Call {CallId} is established ✅", callId);
                    await HandleEstablishedCall(callId);
                    break;

                case "terminated":
                    _logger.LogInformation("[Callback] Call {CallId} has terminated ❌", callId);
                    break;

                default:
                    _logger.LogInformation("[Callback] Unhandled state: {State}", state);
                    break;
            }
        }

        // Always return 200 OK — Graph will retry if you don't
        return Ok();
    }

    // ─── Handle incoming call (bot receives a call) ───────────────────────────

    private async Task HandleIncomingCall(string callId, string? direction)
    {
        // Read behavior from config: "answer" or "reject"
        var behavior = _configuration["Graph:IncomingCallBehavior"] ?? "answer";

        _logger.LogInformation("[Callback] Incoming call {CallId} — behavior: {Behavior}", callId, behavior);

        if (behavior == "answer")
        {
            await _graphService.AnswerCallAsync(callId);
        }
        else
        {
           await _graphService.RejectCallAsync(callId);
        }
    }

    // ─── Handle established call ──────────────────────────────────────────────

    private async Task HandleEstablishedCall(string callId)
    {
        // Auto hang up after call is established (optional)
        var autoHangUp = _configuration.GetValue<bool>("Graph:AutoHangUpOnEstablished");

        if (autoHangUp)
        {
            _logger.LogInformation("[Callback] Auto hang up enabled — ending call {CallId}", callId);
            await Task.Delay(3000); // Wait 3 seconds then hang up
            await _graphService.HangUpCallAsync(callId);
        }
    }
}