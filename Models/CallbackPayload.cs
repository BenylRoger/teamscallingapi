namespace TeamsCallApi.Models;

public class CallbackPayload
{
    public List<CallNotification>? Value { get; set; }
}

public class CallNotification
{
    public string? ChangeType { get; set; }
    public string? Resource { get; set; }
    public CallResource? ResourceData { get; set; }
}

public class CallResource
{
    public string? Id { get; set; }
    public string? State { get; set; }
    public string? Direction { get; set; }
    public string? ResultInfo { get; set; }

    // OData type field from Graph
    public string? OdataType { get; set; }
}