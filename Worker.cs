using System.Text.Json;
using GeneradorDocumentosSQL.Services;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Configuration;
public class Worker : BackgroundService
{
    private readonly DatabaseService _db;
    private readonly WordService _word;
    private readonly PrintService _print;
    private readonly ILogger<Worker> _logger;

    private int _ultimoId = 0;
    private readonly string _estadoFile;
    private readonly string _rutaDestino;
    private readonly string _rutaPlantilla;
    private readonly string _nombreImpresora;
    private readonly int _diasRetencion;
    private readonly int _intervaloMs;

    public Worker(
        IConfiguration config,
        DatabaseService db,
        WordService word,
        PrintService print,
        ILogger<Worker> logger)
    {
        _db = db;
        _word = word;
        _print = print;
        _logger = logger;

        _rutaDestino = Environment.ExpandEnvironmentVariables(
            config["WorkerSettings:OutputPath"] ?? string.Empty
        );
        _rutaPlantilla = Environment.ExpandEnvironmentVariables(
            config["WorkerSettings:TemplatePath"] ?? string.Empty
        );
        _nombreImpresora = config["WorkerSettings:PrinterName"] ?? string.Empty;
        _estadoFile = Path.Combine(AppContext.BaseDirectory, "estado_vta.json");

        _diasRetencion = int.TryParse(config["WorkerSettings:RetentionDays"], out var dias) ? dias : 30;
        _intervaloMs = int.TryParse(config["WorkerSettings:IntervalMs"], out var ms) ? ms : 10000;
    }

    // Carga el último ID procesado desde el archivo de estado
    private void CargarEstado()
    {
        try
        {
            if (File.Exists(_estadoFile))
            {
                var json = File.ReadAllText(_estadoFile);
                var estado = JsonSerializer.Deserialize<EstadoWorker>(json);
                _ultimoId = estado?.UltimoId ?? 0;
                _logger.LogInformation("Estado cargado. UltimoId: {UltimoId}", _ultimoId);
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "No se pudo cargar el estado. Se inicia desde Id 0.");
        }
    }

    private void GuardarEstado()
    {
        try
        {
            File.WriteAllText(_estadoFile, JsonSerializer.Serialize(new EstadoWorker { UltimoId = _ultimoId }));
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error al guardar el estado del worker.");
        }
    }

    private void LimpiarArchivosAntiguos()
    {
        try
        {
            if (!Directory.Exists(_rutaDestino)) return;

            var limite = DateTime.Now.AddDays(-_diasRetencion);
            foreach (var archivo in Directory.GetFiles(_rutaDestino, "*.docx"))
            {
                if (new FileInfo(archivo).CreationTime < limite)
                {
                    File.Delete(archivo);
                    _logger.LogInformation("Archivo eliminado: {Archivo}", archivo);
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error al limpiar archivos antiguos.");
        }
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        CargarEstado();
        LimpiarArchivosAntiguos();

        _logger.LogInformation("Worker iniciado. Escuchando desde IdTicket > {UltimoId}", _ultimoId);

        while (!stoppingToken.IsCancellationRequested)
        {

            try
            {
                var datos = await _db.ObtenerDatosParaCartaAsync(_ultimoId);

                if (datos != null)
                {
                    string rutaTemp = Path.Combine(Path.GetTempPath(), $"Carta_{datos.Folio}_{Guid.NewGuid()}.docx");
                    try
                    {
                        string rutaDoc = await _word.GenerarCartaAsync(datos, _rutaPlantilla, rutaTemp);
                        await _print.ImprimirAsync(rutaDoc, _nombreImpresora);


                        // Espera a que la impresora termine de procesar el documento antes de eliminar el archivo temporal
                        await Task.Delay(20000);

                        _logger.LogInformation(
                        "Carta generada e impresa para cliente: {Nombre}, Ticket: {IdTicket}",
                        datos.Nombre, datos.Folio);

                    }
                    finally
                    {
                        // Se elimina el archivo temporal siempre, incluso si hay error
                        if (File.Exists(rutaTemp))
                        {
                            File.Delete(rutaTemp);
                            _logger.LogInformation("Archivo temporal eliminado: {Ruta}", rutaTemp);
                        }
                    }


                    _ultimoId = datos.Folio;
                    GuardarEstado();


                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error en el ciclo del Worker. Se reintentará en el próximo intervalo.");
            }

            await Task.Delay(_intervaloMs, stoppingToken);
        }
    }
}

// Clase auxiliar para serializar/deserializar el estado
public class EstadoWorker
{
    public int UltimoId { get; set; }
}