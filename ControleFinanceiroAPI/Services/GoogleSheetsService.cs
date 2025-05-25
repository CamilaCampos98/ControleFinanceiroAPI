using ControleFinanceiroAPI.Models;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.IO;

public class GoogleSheetsService
{
    static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
    static readonly string ApplicationName = "ControleFinanceiro";

    private readonly string SpreadsheetId = "16c4P1KwZfuySZ36HSBKvzrl4ZagEXioD6yDhfQ9fhjM";
    private readonly string SheetName = "Controle";

    private readonly SheetsService _service;

    public GoogleSheetsService()
    {
        var jsonFilePath = GetJsonFilePath();
        if (!File.Exists(jsonFilePath))
        {
            throw new FileNotFoundException($"Arquivo de credenciais não encontrado em: {jsonFilePath}");
        }

        GoogleCredential credential;
        using (var stream = new FileStream(jsonFilePath, FileMode.Open, FileAccess.Read))
        {
            credential = GoogleCredential.FromStream(stream)
                                         .CreateScoped(Scopes);
        }

        _service = new SheetsService(new BaseClientService.Initializer
        {
            HttpClientInitializer = credential,
            ApplicationName = ApplicationName
        });
    }

    private static string GetJsonFilePath()
    {
        var pathFromEnv = Environment.GetEnvironmentVariable("GOOGLE_SHEETS_JSON_PATH");

        if (!string.IsNullOrEmpty(pathFromEnv) && File.Exists(pathFromEnv))
        {
            return pathFromEnv;
        }

        // Tenta na raiz
        var localPath = Path.Combine(Directory.GetCurrentDirectory(), "credentials.json");
        if (File.Exists(localPath))
        {
            return localPath;
        }

        // Tenta na pasta wwwroot
        var wwwrootPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "credentials.json");
        if (File.Exists(wwwrootPath))
        {
            return wwwrootPath;
        }

        // Se não encontrar, retorna o primeiro caminho para lançar erro depois
        return pathFromEnv ?? localPath;
    }

    public IList<IList<object>> ReadData(string range)
    {
        var request = _service.Spreadsheets.Values.Get(SpreadsheetId, range);
        var response = request.Execute();
        return response.Values;
    }

    public void WritePurchaseWithInstallments(CompraModel compra)
    {
        var idLan = ObterProximoIdLan();
        compra.idLan = idLan;
        compra.MesAno = compra.Data.ToString("MM/yyyy");

        var linhas = new List<IList<object>>();
        var valorParcela = compra.ValorTotal / compra.TotalParcelas;

        for (int i = 1; i <= compra.TotalParcelas; i++)
        {
            var dataParcela = compra.Data.AddMonths(i - 1);
            var parcelaStr = compra.TotalParcelas > 1 ? $"{i}/{compra.TotalParcelas}" : "";

            linhas.Add(new List<object> {
            compra.idLan,
            compra.FormaPgto,
            parcelaStr,
            compra.Descricao,
            valorParcela.ToString("F2"),
            compra.MesAno,
            dataParcela.ToString("yyyy-MM-dd")
        });
        }

        var valueRange = new ValueRange { Values = linhas };

        var appendRequest = _service.Spreadsheets.Values.Append(
            valueRange,
            SpreadsheetId,
            $"{SheetName}"
        );
        appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;

        var response = appendRequest.Execute();

        // Opcional: log para confirmar inserção
        Console.WriteLine("Linhas inseridas: " + response.Updates.UpdatedRows);
    }


    public void TestConnection()
    {
        var request = _service.Spreadsheets.Get(SpreadsheetId);
        var response = request.Execute();
        Console.WriteLine($"Planilha encontrada: {response.Properties.Title}");
    }

    public int ObterProximoIdLan()
    {
        var range = $"{SheetName}!A:A"; // Supondo que idLan está na coluna A
        var request = _service.Spreadsheets.Values.Get(SpreadsheetId, range);
        var response = request.Execute();

        var valores = response.Values;

        if (valores == null || valores.Count == 0)
            return 1; // Se não tem nenhum, começa com 1

        // Pular o cabeçalho se houver
        var idList = valores.Skip(1) // remove cabeçalho, se não tiver, tira esse skip
                             .Select(l => int.TryParse(l.FirstOrDefault()?.ToString(), out var id) ? id : 0)
                             .Where(id => id > 0)
                             .ToList();

        if (!idList.Any())
            return 1;

        return idList.Max() + 1;
    }

}
