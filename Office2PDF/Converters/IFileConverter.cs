using Microsoft.AspNetCore.Http;

namespace Office2PDF.Converters
{
    internal interface IFileConverter
    {
        bool CanConvert(string extension);
        Task<byte[]> ConvertAsync(IFormFile file, CancellationToken ct = default);
    }
}
