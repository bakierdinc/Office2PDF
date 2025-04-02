using Microsoft.AspNetCore.Http;

namespace Office2PDF.Services
{
    public interface IConversionService
    {
        Task<byte[]> ConvertToPdfAsync(IFormFile file, CancellationToken ct = default);
    }
}