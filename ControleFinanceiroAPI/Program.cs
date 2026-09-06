using System.Globalization;
using ControleFinanceiroAPI;
using ControleFinanceiroAPI.Options;
using ControleFinanceiroAPI.Services;
using Microsoft.OpenApi.Models;

var builder = WebApplication.CreateBuilder(args);

builder.Configuration
    .SetBasePath(Directory.GetCurrentDirectory())
    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
    .AddJsonFile($"appsettings.{builder.Environment.EnvironmentName}.json", optional: true, reloadOnChange: true)
    .AddEnvironmentVariables();



// Add services to the container.
builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen(options =>
{
    options.SwaggerDoc("v1", new OpenApiInfo
    {
        Title = "Controle Financeiro API",
        Version = ApplicationVersion.Current,
        Description = $"Versão do aplicativo: {ApplicationVersion.Current}"
    });
});
builder.Services.Configure<GoogleSheetsOptions>(builder.Configuration.GetSection("GoogleSheets"));
builder.Services.AddSingleton<GoogleSheetsService>();
builder.Services.AddScoped<CompraWorkflowService>();
builder.Services.AddScoped<EntradaWorkflowService>();

builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowAll",
        b => b.AllowAnyOrigin()
                    .AllowAnyMethod()
              .AllowAnyHeader());
});


var app = builder.Build();

if (!app.Environment.IsDevelopment())
{
    var port = Environment.GetEnvironmentVariable("PORT") ?? "10000";
    builder.WebHost.UseUrls($"http://*:{port}");
}

var defaultCulture = new CultureInfo("pt-BR");
CultureInfo.DefaultThreadCurrentCulture = defaultCulture;
CultureInfo.DefaultThreadCurrentUICulture = defaultCulture;

app.UseCors("AllowAll");

app.UseSwagger();
app.UseSwaggerUI(options =>
{
    options.SwaggerEndpoint("/swagger/v1/swagger.json", $"Controle Financeiro API {ApplicationVersion.Current}");
    options.DocumentTitle = $"Controle Financeiro API - {ApplicationVersion.Current}";
});

app.UseHttpsRedirection();

app.UseAuthorization();

app.MapControllers();

app.Run();
