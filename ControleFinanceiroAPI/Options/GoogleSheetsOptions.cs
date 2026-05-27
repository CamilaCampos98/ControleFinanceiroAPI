namespace ControleFinanceiroAPI.Options;

public sealed class GoogleSheetsOptions
{
    public string SpreadsheetId { get; set; } = string.Empty;
    public string SheetName { get; set; } = "Controle";
    public string FixosRange { get; set; } = "Fixos!A:H";
    public string CartoesSheet { get; set; } = "Cartoes";
    public string FixosTipoSheet { get; set; } = "TiposFixos";
}
