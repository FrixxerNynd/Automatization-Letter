using GeneradorDocumentosSQL.Services;

var builder = Host.CreateApplicationBuilder(args);

// Logging explícito
builder.Logging.ClearProviders();
builder.Logging.AddConsole();

// Cadena de conexión con validación
string connectionString = builder.Configuration.GetConnectionString("DefaultConnection")
    ?? throw new InvalidOperationException("No se encontró la cadena de conexión 'DefaultConnection'.");

// DatabaseService necesita ILogger, se resuelve via DI
builder.Services.AddSingleton<DatabaseService>(sp =>
    new DatabaseService(
        connectionString,
        sp.GetRequiredService<ILogger<DatabaseService>>()
    ));

// WordService y PrintService resuelven ILogger automáticamente
builder.Services.AddSingleton<WordService>();
builder.Services.AddSingleton<PrintService>();

builder.Services.AddHostedService<Worker>();

var host = builder.Build();
host.Run();