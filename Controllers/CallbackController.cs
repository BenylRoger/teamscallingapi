using Microsoft.AspNetCore.Mvc;
using TeamsCallApi.Models;

namespace TeamsCallApi.Controllers;

[ApiController]
[Route("api/calls")]
public class CallbackController : ControllerBase
{
    private readonly ILogger<CallbackController> _logger;

    public CallbackController(ILogger<CallbackController> logger)
    {
        _logger = logger;
    }

    /// <summary>
    /// POST /api/calls
    /// Receives callback notifications from Microsoft Graph (e.g., call state changes).
    /// Logs the payload and returns HTTP 200 OK.
    /// </summary>
    [HttpPost]
    public IActionResult HandleCallback([FromBody] CallbackPayload payload)
    {
        _logger.LogInformation("=== [Graph Callback Received] ===");
        _logger.LogInformation("Received at : {Time}", DateTime.UtcNow);
        _logger.LogInformation("Event Type  : {EventType}", payload?.EventType ?? "unknown");
        _logger.LogInformation("Call ID     : {CallId}", payload?.CallId ?? "unknown");
        _logger.LogInformation("State       : {State}", payload?.State ?? "unknown");
        _logger.LogInformation("Raw Payload : {@Payload}", payload);
        _logger.LogInformation("=================================");

        // Microsoft Graph requires HTTP 200 OK to acknowledge callback receipt.
        // If you return anything else, Graph will retry delivery.
        return Ok();
    }
}