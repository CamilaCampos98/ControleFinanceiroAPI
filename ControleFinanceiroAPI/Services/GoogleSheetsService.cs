using ControleFinanceiroAPI.Models;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Microsoft.AspNetCore.Hosting;
using System;
using System.Collections.Generic;
using System.IO;

public class GoogleSheetsService
{
    static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
    static readonly string ApplicationName = "ControleFinanceiro";

    private readonly string SpreadsheetId = "16c4P1KwZfuySZ36HSBKvzrl4ZagEXioD6yDhfQ9fhjM"; // só o ID, sem /edit
    private readonly string SheetName = "Controle";

    static readonly string JsonFilePath = Environment.GetEnvironmentVariable("GOOGLE_SHEETS_JSON_PATH")
                    ?? "C:\\Users\\Camila\\source\\repos\\ControleFinanceiroAPI\\ControleFinanceiroAPI\\wwwroot\\credentials.json";
    private readonly SheetsService _service;

    public SheetsService Connect()
    {
        GoogleCredential credential;
        using (var stream = new FileStream(JsonFilePath, FileMode.Open, FileAccess.Read))
        {
            credential = GoogleCredential.FromStream(stream)
                .CreateScoped(Scopes);
        }

        var service = new SheetsService(new BaseClientService.Initializer()
        {
            HttpClientInitializer = credential,
            ApplicationName = ApplicationName,
        });

        return service;
    }
    public GoogleSheetsService(IWebHostEnvironment env)
    {
        // Ajuste o caminho para o arquivo JSON de credenciais na sua pasta wwwroot
        var jsonPath = Path.Combine(env.WebRootPath, "credentials.json");

        GoogleCredential credential;
        using (var stream = new FileStream(jsonPath, FileMode.Open, FileAccess.Read))
        {
            credential = GoogleCredential.FromStream(stream).CreateScoped(Scopes);
        }

        _service = new SheetsService(new BaseClientService.Initializer()
        {
            HttpClientInitializer = credential,
            ApplicationName = ApplicationName,
        });
    }

  
    // Se quiser, pode deixar o método de leitura aqui também
    public IList<IList<object>> ReadData(string range)
    {
        var request = _service.Spreadsheets.Values.Get(SpreadsheetId, range);
        var response = request.Execute();
        return response.Values;
    }

    public void WritePurchaseWithInstallments(CompraModel compra)
    {
        var service = Connect();

        var linhas = new List<IList<object>>();
        var valorParcela = compra.ValorTotal / compra.TotalParcelas;

        for (int i = 1; i <= compra.TotalParcelas; i++)
        {
            var dataParcela = compra.Data.AddMonths(i - 1);
            var parcelaStr = compra.TotalParcelas > 1 ? $"{i}/{compra.TotalParcelas}" : "";

            linhas.Add(new List<object> {
            compra.FormaPgto,
            parcelaStr,
            compra.Descricao,
            valorParcela.ToString("F2"),
            dataParcela.ToString("yyyy-MM-dd")
        });
        }

        var valueRange = new ValueRange { Values = linhas };

        var appendRequest = service.Spreadsheets.Values.Append(valueRange, SpreadsheetId, $"{SheetName}!A:E");
        appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;

        appendRequest.Execute();
    }


}
