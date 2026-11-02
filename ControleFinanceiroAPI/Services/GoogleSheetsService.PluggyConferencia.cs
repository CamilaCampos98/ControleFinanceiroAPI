using System.Globalization;
using System.Text;

public partial class GoogleSheetsService
{
    public sealed record ParcelaPluggyConferencia(string Id, DateTime Data, decimal Valor, string Pessoa,
        string Cartao, string Parcela, string Descricao);

    public sealed record LancamentoConferencia(int Linha, DateTime Data, decimal Valor, string Pessoa,
        string Cartao, string Parcela, string Descricao);

    public sealed record ResultadoConferenciaPluggy(string Id, string Status,
        LancamentoConferencia? Correspondente, int QuantidadeCorrespondencias);

    public IReadOnlyList<ResultadoConferenciaPluggy> ConferirParcelasPluggy(
        IReadOnlyList<ParcelaPluggyConferencia> parcelas)
    {
        var linhas = ReadData($"{SheetName}!A:J");
        var lancamentos = linhas.Skip(1).Select((linha, indice) => LerLancamento(linha, indice + 2))
            .Where(lancamento => lancamento is not null).Cast<LancamentoConferencia>().ToList();

        return ConferirParcelasPluggy(parcelas, lancamentos);
    }

    public static IReadOnlyList<ResultadoConferenciaPluggy> ConferirParcelasPluggy(
        IReadOnlyList<ParcelaPluggyConferencia> parcelas, IReadOnlyList<LancamentoConferencia> lancamentos)
    {
        // Um lançamento exato pertence à sua própria compra; não pode servir de candidato para outra.
        var linhasExatas = lancamentos.Where(lancamento => parcelas.Any(parcela =>
            Igual(lancamento.Pessoa, parcela.Pessoa) && Igual(lancamento.Cartao, parcela.Cartao) &&
            lancamento.Data == parcela.Data.Date && lancamento.Valor == parcela.Valor))
            .Select(lancamento => lancamento.Linha).ToHashSet();

        return parcelas.Select(parcela => ConferirParcela(parcela, lancamentos, linhasExatas)).ToList();
    }

    private static LancamentoConferencia? LerLancamento(IList<object> linha, int numero)
    {
        var dataTexto = linha.ElementAtOrDefault(6)?.ToString()?.Trim();
        var valorTexto = linha.ElementAtOrDefault(4)?.ToString()?.Trim();
        if (!DateTime.TryParseExact(dataTexto, ["yyyy-MM-dd", "dd/MM/yyyy"],
                CultureInfo.InvariantCulture, DateTimeStyles.None, out var data) ||
            !decimal.TryParse(valorTexto, NumberStyles.Currency, CultureInfo.GetCultureInfo("pt-BR"), out var valor))
            return null;

        return new LancamentoConferencia(numero, data.Date, valor,
            linha.ElementAtOrDefault(7)?.ToString()?.Trim() ?? "",
            linha.ElementAtOrDefault(9)?.ToString()?.Trim() ?? "",
            linha.ElementAtOrDefault(2)?.ToString()?.Trim() ?? "",
            linha.ElementAtOrDefault(3)?.ToString()?.Trim() ?? "");
    }

    private static ResultadoConferenciaPluggy ConferirParcela(ParcelaPluggyConferencia parcela,
        IReadOnlyList<LancamentoConferencia> lancamentos, IReadOnlySet<int> linhasExatas)
    {
        var pessoa = lancamentos.Where(item => Igual(item.Pessoa, parcela.Pessoa)).ToList();
        var exatos = pessoa.Where(item => Igual(item.Cartao, parcela.Cartao) &&
            item.Data == parcela.Data.Date && item.Valor == parcela.Valor).ToList();
        if (exatos.Count > 0)
            return new(parcela.Id, "Existente", exatos[0], exatos.Count);

        var disponiveis = pessoa.Where(item => !linhasExatas.Contains(item.Linha)).ToList();
        var cartaoDivergente = disponiveis.FirstOrDefault(item => !Igual(item.Cartao, parcela.Cartao) &&
            item.Data == parcela.Data.Date && item.Valor == parcela.Valor);
        if (cartaoDivergente is not null)
            return new(parcela.Id, "CartaoDivergente", cartaoDivergente, 1);

        var possivel = disponiveis.Where(item => Igual(item.Cartao, parcela.Cartao))
            .Select(item => new { Item = item, Dias = Math.Abs((item.Data - parcela.Data.Date).Days),
                Diferenca = Math.Abs(item.Valor - parcela.Valor) })
            .Where(candidato => candidato.Dias <= 7 && candidato.Diferenca <= 5m &&
                (candidato.Dias == 0 && MesmaParcela(candidato.Item.Parcela, parcela.Parcela) &&
                    (parcela.Parcela != "1/1" || candidato.Diferenca <= 1m ||
                     DescricaoRelacionada(candidato.Item.Descricao, parcela.Descricao)) ||
                 candidato.Dias == 1 && candidato.Diferenca == 0m ||
                 DescricaoRelacionada(candidato.Item.Descricao, parcela.Descricao) &&
                     MesmaParcela(candidato.Item.Parcela, parcela.Parcela)))
            .OrderBy(candidato => candidato.Dias)
            .ThenBy(candidato => candidato.Diferenca)
            .Select(candidato => candidato.Item).FirstOrDefault();

        return possivel is null
            ? new(parcela.Id, "NaoEncontrada", null, 0)
            : new(parcela.Id, "Revisar", possivel, 1);
    }

    private static bool MesmaParcela(string existente, string pluggy) =>
        string.Equals(string.IsNullOrWhiteSpace(existente) ? "1/1" : existente.Trim(),
            string.IsNullOrWhiteSpace(pluggy) ? "1/1" : pluggy.Trim(), StringComparison.OrdinalIgnoreCase);

    private static bool Igual(string a, string b) =>
        string.Equals(Normalizar(a), Normalizar(b), StringComparison.Ordinal);

    private static bool DescricaoRelacionada(string a, string b)
    {
        var palavras = Normalizar(a).Split(' ', StringSplitOptions.RemoveEmptyEntries)
            .Where(palavra => palavra.Length >= 4).ToHashSet();
        return Normalizar(b).Split(' ', StringSplitOptions.RemoveEmptyEntries)
            .Any(palavra => palavra.Length >= 4 && palavras.Contains(palavra));
    }

    private static string Normalizar(string texto)
    {
        var decomposed = texto.Trim().ToLowerInvariant().Normalize(NormalizationForm.FormD);
        var result = new StringBuilder(decomposed.Length);
        foreach (var character in decomposed)
        {
            if (CharUnicodeInfo.GetUnicodeCategory(character) == UnicodeCategory.NonSpacingMark) continue;
            result.Append(char.IsLetterOrDigit(character) ? character : ' ');
        }
        return result.ToString().Normalize(NormalizationForm.FormC).Trim();
    }
}
