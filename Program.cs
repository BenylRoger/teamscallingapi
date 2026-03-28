using TeamsCallApi.Services;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllers();

// Register GraphService for dependency injection
builder.Services.AddSingleton<GraphService>();

builder.Logging.ClearProviders();
builder.Logging.AddConsole();

var app = builder.Build();

app.UseRouting();
app.MapControllers();

app.Run();