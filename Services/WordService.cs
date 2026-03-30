using GeneradorDocumentosSQL.Models;
using Microsoft.Extensions.Logging;
using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using Xceed.Pdf;

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

                    void ReemplazarParrafo(IEnumerable<Paragraph> parrafos)
                    {
                        foreach (var parrafo in parrafos)
                        {
                            string textoCompleto = string.Concat(parrafo.Descendants<Run>()
                                .Select(r => r.InnerText));

                            if (!textoCompleto.Contains('{')) continue;

                            textoCompleto = textoCompleto
                                .Replace("{nombre_cliente}", cliente.Nombre ?? "N/A")
                                .Replace("{fecha}", fechaFormateada)
                                .Replace("{folio}", cliente.Folio.ToString())
                                .Replace("{fecha_impresora}", fechaImpresora);

                            var primerRun = parrafo.Descendants<Run>().FirstOrDefault();
                            if (primerRun == null) continue;

                            foreach (var run in parrafo.Descendants<Run>().ToList())
                                run.Remove();

                            primerRun.AppendChild(new Text(textoCompleto)
                            {
                                Space = SpaceProcessingModeValues.Preserve
                            });
                            parrafo.AppendChild(primerRun);
                        }
                    }

                    ReemplazarParrafo(body.Descendants<Paragraph>());

                    foreach (var footer in mainPart.FooterParts)
                    {
                        var footerBody = footer.Footer
                        ?? throw new InvalidOperationException("El pie de página no tiene contenido.");

                        ReemplazarParrafo(footerBody.Descendants<Paragraph>());
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
