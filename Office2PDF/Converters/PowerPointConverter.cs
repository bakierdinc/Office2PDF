using Microsoft.AspNetCore.Http;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Office2PDF.Converters.Helpers;

namespace Office2PDF.Converters
{
    internal sealed class PowerPointConverter : IFileConverter
    {
        public bool CanConvert(string extension) => ".pptx".Equals(extension, StringComparison.OrdinalIgnoreCase) || ".ppt".Equals(extension, StringComparison.OrdinalIgnoreCase);

        public async Task<byte[]> ConvertAsync(IFormFile file, CancellationToken ct)
        {
            var tempFolder = FileHelper.GetTempPath();
            var tempInput = await FileHelper.SaveToTempAsync(file, tempFolder, ct);
            var tempOutput = Path.ChangeExtension(tempInput, ".pdf");

            var pptApp = new Application();

            try
            {
                var presentation = pptApp.Presentations.Open(
                    tempInput,
                    ReadOnly: MsoTriState.msoTrue,
                    Untitled: MsoTriState.msoFalse,
                    WithWindow: MsoTriState.msoFalse
                );

                presentation.SaveAs(tempOutput, PpSaveAsFileType.ppSaveAsPDF);
                presentation.Close();

                return await File.ReadAllBytesAsync(tempOutput, ct);
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                FileHelper.SafeDelete(tempInput, tempOutput);
                pptApp.Quit();
            }
        }
    }
}
