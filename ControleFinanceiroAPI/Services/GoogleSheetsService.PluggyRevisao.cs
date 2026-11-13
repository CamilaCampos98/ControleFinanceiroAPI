using System.Globalization;
using Google.Apis.Sheets.v4.Data;

public partial class GoogleSheetsService
{
    public sealed record CorrigirRevisaoPluggyRequest(
        int Linha, string IdLancamento, DateTime DataOriginal, decimal ValorOriginal,
        string PessoaOriginal, string CartaoOriginal, string ParcelaOriginal, string DescricaoOriginal,
        string Compra, decimal Valor, DateTime Data, string Pessoa, string Cartao,
        string FormaPgto, string Fonte);

    public sealed record ParcelaRevisaoPlanejada(LancamentoConferencia Original, DateTime NovaData);

    public int CorrigirRevisaoPluggy(CorrigirRevisaoPluggyRequest request)
    {
        lock (_lockId)
        {
            var rows = ReadData($"{SheetName}!A:J");
            var lancamentos = rows.Skip(1).Select((row, index) => LerLancamento(row, index + 2))
                .Where(item => item is not null).Cast<LancamentoConferencia>().ToList();
            var plano = PlanejarRevisaoPluggy(request, lancamentos);
            var tab = _service.Spreadsheets.Get(SpreadsheetId).Execute().Sheets
                .FirstOrDefault(sheet => sheet.Properties.Title == SheetName)?.Properties.SheetId
                ?? throw new InvalidOperationException("A aba de compras não foi encontrada.");

            var updates = plano.Select(item =>
            {
                var mes = CalcularMesFatura(item.NovaData, request.Cartao.Trim(), request.Pessoa.Trim());
                var values = new ExtendedValue[]
                {
                    TextValue(item.Original.IdLancamento), TextValue(request.FormaPgto.Trim()),
                    TextValue(item.Original.Parcela), TextValue(request.Compra.Trim()),
                    new() { NumberValue = (double)request.Valor }, TextValue(mes),
                    TextValue(item.NovaData.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)),
                    TextValue(request.Pessoa.Trim()), TextValue(request.Fonte.Trim()),
                    TextValue(request.Cartao.Trim())
                };
                return new Request
                {
                    UpdateCells = new UpdateCellsRequest
                    {
                        Range = new GridRange
                        {
                            SheetId = tab,
                            StartRowIndex = item.Original.Linha - 1,
                            EndRowIndex = item.Original.Linha,
                            StartColumnIndex = 0,
                            EndColumnIndex = 10
                        },
                        Rows = [new RowData { Values = values.Select(value => new CellData
                        { UserEnteredValue = value }).ToList() }],
                        Fields = "userEnteredValue"
                    }
                };
            }).ToList();

            _service.Spreadsheets.BatchUpdate(new BatchUpdateSpreadsheetRequest
            { Requests = updates }, SpreadsheetId).Execute();
            return updates.Count;
        }
    }

    public static IReadOnlyList<ParcelaRevisaoPlanejada> PlanejarRevisaoPluggy(
        CorrigirRevisaoPluggyRequest request, IReadOnlyList<LancamentoConferencia> lancamentos)
    {
        if (request.Linha < 2 || string.IsNullOrWhiteSpace(request.IdLancamento) ||
            string.IsNullOrWhiteSpace(request.PessoaOriginal) ||
            string.IsNullOrWhiteSpace(request.CartaoOriginal) ||
            request.DescricaoOriginal is null || request.ParcelaOriginal is null ||
            string.IsNullOrWhiteSpace(request.Compra) || request.Valor <= 0 ||
            decimal.Round(request.Valor, 2) != request.Valor ||
            request.Data.Year is < 2000 or > 2100 ||
            string.IsNullOrWhiteSpace(request.Pessoa) ||
            request.Pessoa.Trim() is not ("Camila" or "Diego") ||
            string.IsNullOrWhiteSpace(request.Cartao) ||
            string.IsNullOrWhiteSpace(request.FormaPgto) ||
            string.IsNullOrWhiteSpace(request.Fonte))
            throw new ArgumentException("Preencha todos os dados da compra com valores válidos.");

        var escolhida = lancamentos.SingleOrDefault(item => item.Linha == request.Linha);
        if (escolhida is null || escolhida.IdLancamento != request.IdLancamento ||
            escolhida.Data.Date != request.DataOriginal.Date ||
            escolhida.Valor != request.ValorOriginal ||
            !Igual(escolhida.Pessoa, request.PessoaOriginal) ||
            !Igual(escolhida.Cartao, request.CartaoOriginal) ||
            !MesmaParcela(escolhida.Parcela, request.ParcelaOriginal) ||
            !string.Equals(escolhida.Descricao, request.DescricaoOriginal.Trim(), StringComparison.Ordinal))
            throw new InvalidOperationException("A linha da planilha mudou. Recarregue a conferência antes de salvar.");

        var grupo = lancamentos.Where(item => item.IdLancamento == escolhida.IdLancamento).ToList();
        if (!TryParseParcela(escolhida.Parcela, out var numeroEscolhido, out var total))
        {
            if (grupo.Count != 1)
                throw new InvalidOperationException("O ID da compra aparece em mais de uma linha sem parcelas válidas.");
            return [new ParcelaRevisaoPlanejada(escolhida, request.Data.Date)];
        }

        if (grupo.Count != total || !Enumerable.Range(1, total).All(numero => grupo.Count(item =>
                MesmaParcela(item.Parcela, $"{numero}/{total}") &&
                Igual(item.Pessoa, escolhida.Pessoa) && Igual(item.Cartao, escolhida.Cartao)) == 1))
            throw new InvalidOperationException("A série de parcelas está incompleta ou inconsistente. Nenhuma linha foi alterada.");

        return grupo.Select(item =>
        {
            TryParseParcela(item.Parcela, out var numero, out _);
            return new ParcelaRevisaoPlanejada(item, request.Data.Date.AddMonths(numero - numeroEscolhido));
        }).OrderBy(item => item.Original.Linha).ToList();
    }

    private static ExtendedValue TextValue(string value) => new() { StringValue = value };
}
