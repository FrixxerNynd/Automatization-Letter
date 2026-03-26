using System;
using System.Diagnostics;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;

namespace GeneradorDocumentosSQL.Services
{
    public class PrintService
    {
        private readonly ILogger<PrintService> _logger;

        public PrintService(ILogger<PrintService> logger)
        {
            _logger = logger;
        }

        public async Task ImprimirAsync(string rutaDocumento, string impresora)
        {
            if (string.IsNullOrWhiteSpace(rutaDocumento))
                throw new ArgumentException("La ruta del documento no puede estar vacía.", nameof(rutaDocumento));

            if (string.IsNullOrWhiteSpace(impresora))
                throw new ArgumentException("El nombre de la impresora no puede estar vacío.", nameof(impresora));

            try
            {
                _logger.LogInformation("Iniciando impresión de '{Documento}' en '{Impresora}'", rutaDocumento, impresora);

                var info = new ProcessStartInfo
                {
                    FileName = rutaDocumento,
                    Verb = "printto",
                    Arguments = $"\"{impresora}\"",
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden,
                    UseShellExecute = true // Necesario para que funcione el Verb "printto"
                };

                var process = Process.Start(info)
                    ?? throw new InvalidOperationException("No se pudo iniciar el proceso de impresión.");

                await process.WaitForExitAsync();

                if (process.ExitCode != 0)
                    _logger.LogWarning("El proceso de impresión terminó con código de salida {Codigo}", process.ExitCode);
                else
                    _logger.LogInformation("Impresión completada exitosamente.");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al imprimir '{Documento}' en '{Impresora}'", rutaDocumento, impresora);
                throw;
            }
        }
    }
}