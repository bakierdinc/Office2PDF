using Microsoft.AspNetCore.Http;

namespace Office2PDF.Converters.Helpers
{
    internal static class FileHelper
    {
        public static string GetTempPath()
        {
            return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "temp");
        }

        public static async Task<string> SaveToTempAsync(IFormFile file, string tempFolder, CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();

            Directory.CreateDirectory(tempFolder);
            var tempFile = Path.Combine(tempFolder, $"{Guid.NewGuid()}{Path.GetExtension(file.FileName)}");

            await using var stream = File.Create(tempFile);
            await file.CopyToAsync(stream, ct);

            return tempFile;
        }

        public static void SafeDelete(params string[] files)
        {
            if (files is null)
            {
                return;
            }

            foreach (var file in files)
            {
                if (File.Exists(file))
                {
                    try
                    {
                        File.Delete(file);
                    }
                    catch
                    {
                    }
                }
            }
        }
    }
}
