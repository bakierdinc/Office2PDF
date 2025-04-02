using System.Text.Json;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Diagnostics;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Routing;
using Office2PDF;
using Office2PDF.Converters;
using Office2PDF.Services;
using Serilog;

var serviceName = Meta.Assembly.GetName().Name;

if (WindowsServiceInstaller.SetupIfRequired(args, serviceName) is SetupResult.Setup)
{
    return;
}

var builder = WebApplication.CreateBuilder(args);

var logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs", ".log");

Log.Logger = new LoggerConfiguration()
    .WriteTo.File(logPath, rollingInterval: RollingInterval.Day)
    .CreateLogger();

try
{
    builder.Services.AddControllers();
    builder.Services.AddEndpointsApiExplorer();
    builder.Services.AddSwaggerGen();
    builder.Services.Configure<RouteOptions>(options =>
    {
        options.LowercaseUrls = true;
        options.LowercaseQueryStrings = true;
    });
    builder.Services.AddWindowsService(options =>
    {
        options.ServiceName = serviceName;
    });
    builder.WebHost.ConfigureKestrel((context, options) =>
    {
        options.Configure(context.Configuration.GetSection("Kestrel"));
    });
    builder.Services.AddSingleton<IConversionService, ConversionService>();
    builder.Services.AddSingleton<IFileConverter, ExcelConverter>();
    builder.Services.AddSingleton<IFileConverter, WordConverter>();
    builder.Services.AddSingleton<IFileConverter, PowerPointConverter>();

    builder.Logging.ClearProviders();
    builder.Logging.AddSerilog();

    var app = builder.Build();
    app.UseSwagger();
    app.UseSwaggerUI(options =>
    {
        options.SwaggerEndpoint("/swagger/v1/swagger.json", $"{serviceName} API");
        options.RoutePrefix = string.Empty;
    });
    app.MapControllers();
    app.UseExceptionHandler(errorApp =>
    {
        errorApp.Run(async context =>
        {
            var exceptionHandlerPathFeature = context.Features.Get<IExceptionHandlerPathFeature>();
            var exception = exceptionHandlerPathFeature?.Error;

            context.Response.StatusCode = 500;
            context.Response.ContentType = "application/json";

            await context.Response.WriteAsync(JsonSerializer.Serialize(new
            {
                error = "An unexpected error occurred.",
                detail = exception.Message
            }));
        });
    });
    app.Run();
}
catch (Exception e)
{
    Log.Logger.Error(e, e.Message);
    throw;
}
finally
{
    Log.CloseAndFlush();
}