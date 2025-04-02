using Microsoft.AspNetCore.Http;
using Microsoft.Office.Interop.Excel;
using Office2PDF.Converters.Helpers;

namespace Office2PDF.Converters
{
    internal sealed class ExcelConverter : IFileConverter
    {
        public bool CanConvert(string extension) => extension == ".xlsx" || extension == ".xls";

        public async Task<byte[]> ConvertAsync(IFormFile file, CancellationToken ct)
        {
            var tempFolder = FileHelper.GetTempPath();
            var tempInput = await FileHelper.SaveToTempAsync(file, tempFolder, ct);
            var tempOutput = Path.ChangeExtension(tempInput, ".pdf");

            var excelApp = new Application
            {
                DisplayAlerts = false,
                Visible = false,
            };

            try
            {
                var workbook = excelApp.Workbooks.Open(tempInput);
                workbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, tempOutput);
                workbook.Close(false);

                return await File.ReadAllBytesAsync(tempOutput, ct);
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                FileHelper.SafeDelete(tempInput, tempOutput);
                excelApp.Quit();
            }
        }
    }
}