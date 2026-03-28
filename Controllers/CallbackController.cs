using Microsoft.AspNetCore.Mvc;
using System.Text.Json;
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
        _logger       = logger;
        _graphService = graphService;
        _configuration = configuration;
    }

    [HttpPost]
    public async Task<IActionResult> HandleCallback()
    {
        // Read raw body
        using var reader = new StreamReader(Request.Body);
        var rawBody = await reader.ReadToEndAsync();

        _logger.LogInformation("=== [Graph Callback Raw] ===");
        _logger.LogInformation("{RawBody}", rawBody);
        _logger.LogInformation("============================");

        try
        {
            var json = JsonDocument.Parse(rawBody);
            var root = json.RootElement;

            // Graph sends { "value": [ { "resourceData": { "id": "", "state": "" } } ] }
            if (root.TryGetProperty("value", out var valueArray))
            {
                foreach (var notification in valueArray.EnumerateArray())
                {
                    string? callId = null;
                    string? state  = null;

                    // ✅ Extract callId from resource URL path
                    // e.g. "/app/calls/01004980-be37-40a2-a500-15ed0b24a585"
                    if (notification.TryGetProperty("resource", out var resource))
                    {
                        var resourcePath = resource.GetString();
                        callId = resourcePath?.Split('/').LastOrDefault();
                    }

                    if (notification.TryGetProperty("resourceData", out var resourceData))
                    {
                        state = resourceData.TryGetProperty("state", out var stateVal)
                            ? stateVal.GetString()?.ToLower() : null;
                    }

                    _logger.LogInformation("[Callback] CallId: {CallId} | State: {State}", callId, state);

                    if (string.IsNullOrEmpty(callId)) continue;

                    switch (state)
                    {
                        case "incoming":
                            await HandleIncomingCall(callId);
                            break;

                        case "establishing":
                            _logger.LogInformation("[Callback] Call {CallId} establishing...", callId);
                            break;

                        case "established":
                            // Only trigger once when audio becomes active
                            if (resourceData.TryGetProperty("mediaState", out var mediaState))
                            {
                                if (mediaState.TryGetProperty("audio", out var audio) &&
                                    audio.GetString() == "active")
                                {
                                    _logger.LogInformation("[Callback] Audio active on call {CallId} ✅ — playing prompt", callId);
                                    await HandleEstablishedCall(callId);
                                }
                            }
                            else
                            {
                                _logger.LogInformation("[Callback] Call {CallId} established — waiting for audio active", callId);
                            }
                            break;

                        case "terminated":
                            _logger.LogInformation("[Callback] Call {CallId} terminated ❌", callId);
                            break;

                        default:
                            _logger.LogInformation("[Callback] Unknown state: {State}", state);
                            break;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogError("[Callback] Failed to parse payload: {Message}", ex.Message);
        }

        // Always return 200 OK
        return Ok();
    }

    private async Task HandleIncomingCall(string callId)
    {
        var behavior = _configuration["Graph:IncomingCallBehavior"] ?? "answer";
        _logger.LogInformation("[Callback] Incoming call — behavior: {Behavior}", behavior);

        if (behavior == "answer")
            await _graphService.AnswerCallAsync(callId);
        else
            await _graphService.RejectCallAsync(callId);
    }

    private async Task HandleEstablishedCall(string callId)
    {
        var audioUrl = _configuration["Graph:AudioPromptUrl"];

        if (!string.IsNullOrEmpty(audioUrl))
        {
            try
            {
                await _graphService.PlayPromptAsync(callId, audioUrl);

                var waitSeconds = _configuration.GetValue<int>("Graph:AudioWaitSeconds", 10);
                _logger.LogInformation("[Callback] Waiting {Seconds}s for audio to finish...", waitSeconds);
                await Task.Delay(TimeSpan.FromSeconds(waitSeconds));
            }
            catch (Exception ex)
            {
                _logger.LogError("[Callback] Audio play failed: {Message}", ex.Message);
            }
        }

        _logger.LogInformation("[Callback] Ending call {CallId}", callId);
        await _graphService.HangUpCallAsync(callId);
    }
}