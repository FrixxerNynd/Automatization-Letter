using System;
using System.Diagnostics;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using System.Management;

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
                    WindowStyle = ProcessWindowStyle.Normal,
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

        private void LimpiarColaImpresora(string impresora)
        {
            try
            {
                using var searcher = new ManagementObjectSearcher($"SELECT * FROM Win32_PrintJob WHERE Name LIKE '{impresora}%'");

                var trabajos = searcher.Get();

                if (!trabajos.Count.Equals(0))
                {
                    _logger.LogWarning("Se encontraron {Count} trabajos atorados en la cola. Limpiando...",
                        trabajos.Count);

                    foreach (ManagementObject trabajo in trabajos)
                        trabajo.Delete();

                    _logger.LogInformation("Cola de impresión limpiada.");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al limpiar la cola de impresión.");
            }
        }
    }
}
