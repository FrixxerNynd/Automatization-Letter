using GeneradorDocumentosSQL.Models;
using Microsoft.Extensions.Logging;
using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace GeneradorDocumentosSQL.Services
{
    public class WordService
    {
        private readonly ILogger<WordService> _logger;

        public WordService(ILogger<WordService> logger)
        {
            _logger = logger;
        }

        public async Task<string> GenerarCartaAsync(Cliente cliente, string templatePath, string outputPath)
        {
            if (!File.Exists(templatePath))
                throw new FileNotFoundException("No se encontró la plantilla Word.", templatePath);

            string nombreLimpio = string.Concat(
                (cliente.Nombre ?? "SinNombre").Split(Path.GetInvalidFileNameChars())
            );
            string nombreArchivo = $"Carta_{cliente.Folio}_{nombreLimpio}.docx";
            string rutaSalida = Path.Combine(outputPath, nombreArchivo);

            if (!Directory.Exists(outputPath))
                Directory.CreateDirectory(outputPath);

            var culturaEs = new CultureInfo("es-MX");
            string fechaFormateada = DateTime.Now.ToString("dd 'de' MMMM 'de' yyyy", culturaEs);

            await Task.Run(() =>
            {
                File.Copy(templatePath, rutaSalida, overwrite: true);

                using var doc = WordprocessingDocument.Open(rutaSalida, isEditable: true);

                var mainPart = doc.MainDocumentPart
                    ?? throw new InvalidOperationException("El documento no tiene parte principal.");

                var body = mainPart.Document?.Body
                    ?? throw new InvalidOperationException("El documento no tiene cuerpo.");

                // Recorre cada párrafo y consolida el texto antes de reemplazar
                foreach (var parrafo in body.Descendants<Paragraph>())
                {
                    // Obtiene el texto completo del párrafo uniendo todos los runs
                    string textoCompleto = string.Concat(parrafo.Descendants<Run>()
                        .Select(r => r.InnerText));

                    // Solo procesa párrafos que tengan placeholders
                    if (!textoCompleto.Contains('{')) continue;

                    textoCompleto = textoCompleto
                        .Replace("{nombre_cliente}", cliente.Nombre ?? "N/A")
                        .Replace("{fecha}", fechaFormateada)
                        .Replace("{folio}", cliente.Folio.ToString());

                    // Limpia los runs existentes y deja solo uno con el texto reemplazado
                    var primerRun = parrafo.Descendants<Run>().FirstOrDefault();
                    if (primerRun == null) continue;

                    // Elimina todos los runs del párrafo
                    foreach (var run in parrafo.Descendants<Run>().ToList())
                        run.Remove();

                    // Inserta un run limpio con el texto ya reemplazado
                    primerRun.GetFirstChild<Text>()?.Remove();
                    primerRun.AppendChild(new Text(textoCompleto)
                    {
                        Space = SpaceProcessingModeValues.Preserve
                    });
                    parrafo.AppendChild(primerRun);
                }

                mainPart.Document.Save();
            });


            _logger.LogInformation("Carta generada: {RutaSalida}", rutaSalida);
            return rutaSalida;
        }
    }
}
