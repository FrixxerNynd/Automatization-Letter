using GeneradorDocumentosSQL.Models;
using Microsoft.Extensions.Logging;
using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System;
using System.IO;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Linq;

namespace GeneradorDocumentosSQL.Services
{
    public class WordService
    {
        private readonly ILogger<WordService> _logger;

        public WordService(ILogger<WordService> logger)
        {
            _logger = logger;
        }

        public async Task<string> GenerarCartaAsync(Cliente cliente, string templatePath, string rutaSalida)
        {
            if (!File.Exists(templatePath))
                throw new FileNotFoundException("No se encontró la plantilla Word.", templatePath);

            try
            {
                var culturaEs = new CultureInfo("es-MX");
                string fechaFormateada = DateTime.Now.ToString("dd 'de' MMMM 'de' yyyy", culturaEs);
                string fechaImpresora = DateTime.Now.ToString("dd 'de' MMMM 'de' yyyy 'a las' HH:mm", culturaEs);

                await Task.Run(() =>
                {
                    File.Copy(templatePath, rutaSalida, overwrite: true);

                    using var doc = WordprocessingDocument.Open(rutaSalida, isEditable: true);

                    var mainPart = doc.MainDocumentPart
                        ?? throw new InvalidOperationException("El documento no tiene parte principal.");

                    var body = mainPart.Document?.Body
                        ?? throw new InvalidOperationException("El documento no tiene cuerpo.");

                    void ReemplazarEnParrafos(IEnumerable<Paragraph> parrafos)
                    {
                        foreach (var parrafo in parrafos.ToList())
                        {
                            // Verifica si el párrafo tiene algún placeholder
                            string textoParrafo = string.Concat(
                                parrafo.Descendants<Run>().Select(r => r.InnerText));

                            if (!textoParrafo.Contains('{')) continue;

                            // Reemplaza en cada run individualmente conservando su formato
                            foreach (var run in parrafo.Descendants<Run>().ToList())
                            {
                                var textoElement = run.GetFirstChild<Text>();
                                if (textoElement == null) continue;

                                string textoRun = textoElement.Text;
                                if (!textoRun.Contains('{')) continue;

                                textoElement.Text = textoRun
                                    .Replace("{nombre_cliente}", cliente.Nombre ?? "N/A")
                                    .Replace("{fecha}", fechaFormateada)
                                    .Replace("{folio}", cliente.Folio.ToString())
                                    .Replace("{fecha_impresion}", fechaImpresora);

                                textoElement.Space = SpaceProcessingModeValues.Preserve;
                            }
                        }
                    }

                    // Reemplaza en el cuerpo del documento
                    ReemplazarEnParrafos(body.Descendants<Paragraph>());

                    // Reemplaza en todos los pies de página
                    foreach (var footer in mainPart.FooterParts)
                    {
                        var footerBody = footer.Footer
                            ?? throw new InvalidOperationException("El pie de página no tiene contenido.");

                        ReemplazarEnParrafos(footerBody.Descendants<Paragraph>());
                        footer.Footer.Save();
                    }

                    mainPart.Document.Save();
                });

                _logger.LogInformation("Carta generada: {RutaSalida}", rutaSalida);
                return rutaSalida;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al generar carta para cliente {IdCliente}, folio {Folio}",
                    cliente.IdCliente, cliente.Folio);
                throw;
            }
        }
    }
}