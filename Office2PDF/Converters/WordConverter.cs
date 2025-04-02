using Microsoft.AspNetCore.Http;
using Microsoft.Office.Interop.Word;
using Office2PDF.Converters.Helpers;

namespace Office2PDF.Converters
{
    internal sealed class WordConverter : IFileConverter
    {
        public bool CanConvert(string extension) => ".docx".Equals(extension, StringComparison.OrdinalIgnoreCase);

        public async Task<byte[]> ConvertAsync(IFormFile file, CancellationToken ct)
        {
            var tempFolder = FileHelper.GetTempPath();
            var tempInput = await FileHelper.SaveToTempAsync(file, tempFolder, ct);
            var tempOutput = Path.ChangeExtension(tempInput, ".pdf");

            var wordApp = new Application();

            try
            {
                var doc = wordApp.Documents.Open(tempInput, ReadOnly: true, Visible: false);
                doc.ExportAsFixedFormat(tempOutput, WdExportFormat.wdExportFormatPDF);
                doc.Close();

                return await File.ReadAllBytesAsync(tempOutput, ct);
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                FileHelper.SafeDelete(tempInput, tempOutput);
                wordApp.Quit();
            }
        }
    }
}
