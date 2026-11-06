using System.Globalization;
using System.Text;

public partial class GoogleSheetsService
{
    public sealed record ParcelaPluggyConferencia(string Id, DateTime Data, decimal Valor, string Pessoa,
        string Cartao, string Parcela, string Descricao, DateTime? DataInformadaPluggy = null);

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
            CorrespondenciaExata(lancamento, parcela)))
            .Select(lancamento => lancamento.Linha).ToHashSet();

        var candidatosPorParcela = parcelas.Select(parcela => lancamentos.Any(item =>
                CorrespondenciaExata(item, parcela))
            ? new List<LancamentoConferencia>()
            : lancamentos.Where(item => !linhasExatas.Contains(item.Linha) &&
                DataProximaDaMesmaCompra(item, parcela)).ToList()).ToList();
        var disputasPorLinha = candidatosPorParcela.SelectMany(candidatos => candidatos)
            .GroupBy(item => item.Linha).ToDictionary(grupo => grupo.Key, grupo => grupo.Count());

        return parcelas.Select((parcela, indice) => ConferirParcela(parcela, lancamentos, linhasExatas,
            candidatosPorParcela[indice], disputasPorLinha)).ToList();
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
        IReadOnlyList<LancamentoConferencia> lancamentos, IReadOnlySet<int> linhasExatas,
        IReadOnlyList<LancamentoConferencia> candidatosProximos, IReadOnlyDictionary<int, int> disputasPorLinha)
    {
        var pessoa = lancamentos.Where(item => Igual(item.Pessoa, parcela.Pessoa)).ToList();
        var exatos = pessoa.Where(item => CorrespondenciaExata(item, parcela)).ToList();
        if (exatos.Count > 0)
            return new(parcela.Id, "Existente", exatos[0], exatos.Count);

        if (candidatosProximos.Count == 1 && disputasPorLinha[candidatosProximos[0].Linha] == 1)
            return new(parcela.Id, "Existente", candidatosProximos[0], 1);
        if (candidatosProximos.Count > 0)
            return new(parcela.Id, "Revisar", candidatosProximos[0], candidatosProximos.Count);

        var disponiveis = pessoa.Where(item => !linhasExatas.Contains(item.Linha)).ToList();
        var cartaoDivergente = disponiveis.FirstOrDefault(item => !Igual(item.Cartao, parcela.Cartao) &&
            item.Data == parcela.Data.Date && item.Valor == parcela.Valor &&
            (!DataFoiAjustada(parcela) || DescricaoRelacionada(item.Descricao, parcela.Descricao)));
        if (cartaoDivergente is not null)
            return new(parcela.Id, "CartaoDivergente", cartaoDivergente, 1);

        return new(parcela.Id, "NaoEncontrada", null, 0);
    }

    private static bool DataProximaDaMesmaCompra(LancamentoConferencia item, ParcelaPluggyConferencia parcela) =>
        Igual(item.Pessoa, parcela.Pessoa) && Igual(item.Cartao, parcela.Cartao) &&
        item.Valor == parcela.Valor && Math.Abs((item.Data - parcela.Data.Date).Days) == 1 &&
        MesmaParcela(item.Parcela, parcela.Parcela) && DescricaoRelacionada(item.Descricao, parcela.Descricao);

    private static bool CorrespondenciaExata(LancamentoConferencia item, ParcelaPluggyConferencia parcela) =>
        Igual(item.Pessoa, parcela.Pessoa) && Igual(item.Cartao, parcela.Cartao) &&
        item.Data == parcela.Data.Date && item.Valor == parcela.Valor &&
        (!DataFoiAjustada(parcela) || DescricaoRelacionada(item.Descricao, parcela.Descricao));

    private static bool DataFoiAjustada(ParcelaPluggyConferencia parcela) =>
        parcela.DataInformadaPluggy.HasValue && parcela.DataInformadaPluggy.Value.Date != parcela.Data.Date;

    private static bool MesmaParcela(string existente, string pluggy) =>
        string.Equals(string.IsNullOrWhiteSpace(existente) ? "1/1" : existente.Trim(),
            string.IsNullOrWhiteSpace(pluggy) ? "1/1" : pluggy.Trim(), StringComparison.OrdinalIgnoreCase);

    private static bool Igual(string a, string b) =>
        string.Equals(Normalizar(a), Normalizar(b), StringComparison.Ordinal);

    private static bool DescricaoRelacionada(string a, string b)
    {
        var palavras = Normalizar(a).Split(' ', StringSplitOptions.RemoveEmptyEntries)
            .Where(palavra => palavra.Length >= 5).ToHashSet();
        return Normalizar(b).Split(' ', StringSplitOptions.RemoveEmptyEntries)
            .Any(palavra => palavra.Length >= 5 && palavras.Any(outra =>
                outra.StartsWith(palavra, StringComparison.Ordinal) ||
                palavra.StartsWith(outra, StringComparison.Ordinal)));
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
