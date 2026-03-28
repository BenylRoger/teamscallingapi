namespace TeamsCallApi.Models;

/// <summary>
/// Represents the JSON payload sent by Microsoft Graph call notifications.
/// Extend this class with additional fields as needed.
/// </summary>
public class CallbackPayload
{
    public string? EventType { get; set; }
    public string? CallId { get; set; }
    public string? State { get; set; }
    public string? TenantId { get; set; }
    public object? AdditionalData { get; set; }
}