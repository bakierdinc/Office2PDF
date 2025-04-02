using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Office2PDF.Services;

namespace Office2PDF.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ConvertController : ControllerBase
    {
        private readonly IConversionService _conversionService;
        private readonly ILogger<ConvertController> _logger;

        public ConvertController(IConversionService conversionService, ILogger<ConvertController> logger)
        {
            _conversionService = conversionService;
            _logger = logger;
        }

        [HttpPost]
        public async Task<IActionResult> Convert(IFormFile file, CancellationToken ct = default)
        {
            var result = await _conversionService.ConvertToPdfAsync(file, ct);

            if (result is null || result.Length <= 0)
            {
                return new StatusCodeResult(StatusCodes.Status500InternalServerError);
            }

            return File(result, "application/pdf", fileDownloadName: Path.ChangeExtension(file.FileName, ".pdf"));
        }
    }
}