using Microsoft.AspNetCore.Http;
using Office2PDF.Converters;

namespace Office2PDF.Services
{
    internal sealed class ConversionService : IConversionService
    {
        private readonly IEnumerable<IFileConverter> _converters;

        public ConversionService(IEnumerable<IFileConverter> converters)
        {
            _converters = converters;
        }

        public async Task<byte[]> ConvertToPdfAsync(IFormFile file, CancellationToken ct = default)
        {
            var extension = Path.GetExtension(file.FileName).ToLowerInvariant();

            var converter = _converters.FirstOrDefault(p => p.CanConvert(extension));
            if (converter is null)
            {
                throw new NotSupportedException($"Unsupported file extension: {extension}");
            }

            return await converter.ConvertAsync(file, ct);
        }
    }
}