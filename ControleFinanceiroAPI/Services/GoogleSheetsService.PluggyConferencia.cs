using System.Globalization;
using System.Text;

public partial class GoogleSheetsService
{
    public sealed record ParcelaPluggyConferencia(string Id, DateTime Data, decimal Valor, string Pessoa,
        string Cartao, string Parcela, string Descricao, DateTime? DataInformadaPluggy = null);

    public sealed record LancamentoConferencia(int Linha, DateTime Data, decimal Valor, string Pessoa,
        string Cartao, string Parcela, string Descricao, string IdLancamento = "",
        string FormaPgto = "", string Fonte = "", string MesAno = "");

    public sealed record ResultadoConferenciaPluggy(string Id, string Status,
        LancamentoConferencia? Correspondente, int QuantidadeCorrespondencias,
        IReadOnlyList<LancamentoConferencia>? Candidatos = null);

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
            linha.ElementAtOrDefault(3)?.ToString()?.Trim() ?? "",
            linha.ElementAtOrDefault(0)?.ToString()?.Trim() ?? "",
            linha.ElementAtOrDefault(1)?.ToString()?.Trim() ?? "",
            linha.ElementAtOrDefault(8)?.ToString()?.Trim() ?? "",
            linha.ElementAtOrDefault(5)?.ToString()?.Trim() ?? "");
    }

    private static ResultadoConferenciaPluggy ConferirParcela(ParcelaPluggyConferencia parcela,
        IReadOnlyList<LancamentoConferencia> lancamentos, IReadOnlySet<int> linhasExatas,
        IReadOnlyList<LancamentoConferencia> candidatosProximos, IReadOnlyDictionary<int, int> disputasPorLinha)
    {
        var pessoa = lancamentos.Where(item => Igual(item.Pessoa, parcela.Pessoa)).ToList();
        var exatos = pessoa.Where(item => CorrespondenciaExata(item, parcela)).ToList();
        if (exatos.Count > 0)
            return new(parcela.Id, exatos.Count == 1 ? "Existente" : "Revisar", exatos[0], exatos.Count, exatos);

        var iniciosDaSerie = EncontrarIniciosDaSerie(parcela, lancamentos);
        if (iniciosDaSerie.Count > 0)
        {
            var inicio = iniciosDaSerie[0];
            var completa = iniciosDaSerie.Count == 1 &&
                SerieCompleta(inicio, parcela, lancamentos);
            return new(parcela.Id, completa ? "Existente" : "Revisar", inicio, iniciosDaSerie.Count, iniciosDaSerie);
        }

        if (candidatosProximos.Count == 1 && disputasPorLinha[candidatosProximos[0].Linha] == 1)
            return new(parcela.Id, "Existente", candidatosProximos[0], 1, candidatosProximos);
        if (candidatosProximos.Count > 0)
            return new(parcela.Id, "Revisar", candidatosProximos[0], candidatosProximos.Count, candidatosProximos);

        if (TryParseParcela(parcela.Parcela, out var numero, out _) && numero == 1)
        {
            var inicioComValorProximo = pessoa.Where(item => item.IdLancamento.Length > 0 &&
                !linhasExatas.Contains(item.Linha) && Igual(item.Cartao, parcela.Cartao) &&
                MesmaParcela(item.Parcela, parcela.Parcela) &&
                Math.Abs((item.Data - parcela.Data.Date).Days) <= 1 &&
                Math.Abs(item.Valor - parcela.Valor) is > 0 and <= 0.10m).ToList();
            if (inicioComValorProximo.Count > 0)
                return new(parcela.Id, "Revisar", inicioComValorProximo[0], inicioComValorProximo.Count, inicioComValorProximo);
        }

        // Uma data convertida de UTC pode coincidir com uma compra de outro estabelecimento.
        // Não assumir que é a mesma: exigir revisão antes de permitir nova importação.
        if (DataFoiAjustada(parcela))
        {
            var mesmaDataEValor = pessoa.Where(item => !linhasExatas.Contains(item.Linha) &&
                Igual(item.Cartao, parcela.Cartao) && MesmaParcela(item.Parcela, parcela.Parcela) &&
                item.Data == parcela.Data.Date && item.Valor == parcela.Valor).ToList();
            if (mesmaDataEValor.Count > 0)
                return new(parcela.Id, "Revisar", mesmaDataEValor[0], mesmaDataEValor.Count, mesmaDataEValor);
        }

        var disponiveis = pessoa.Where(item => !linhasExatas.Contains(item.Linha)).ToList();
        var cartaoDivergente = disponiveis.FirstOrDefault(item => !Igual(item.Cartao, parcela.Cartao) &&
            item.Data == parcela.Data.Date && item.Valor == parcela.Valor &&
            (!DataFoiAjustada(parcela) || DescricaoRelacionada(item.Descricao, parcela.Descricao)));
        if (cartaoDivergente is not null)
            return new(parcela.Id, "CartaoDivergente", cartaoDivergente, 1, [cartaoDivergente]);

        return new(parcela.Id, "NaoEncontrada", null, 0);
    }

    private static bool DataProximaDaMesmaCompra(LancamentoConferencia item, ParcelaPluggyConferencia parcela) =>
        Igual(item.Pessoa, parcela.Pessoa) && Igual(item.Cartao, parcela.Cartao) &&
        item.Valor == parcela.Valor && Math.Abs((item.Data - parcela.Data.Date).Days) == 1 &&
        MesmaParcela(item.Parcela, parcela.Parcela) && DescricaoRelacionada(item.Descricao, parcela.Descricao);

    private static bool CorrespondenciaExata(LancamentoConferencia item, ParcelaPluggyConferencia parcela) =>
        Igual(item.Pessoa, parcela.Pessoa) && Igual(item.Cartao, parcela.Cartao) &&
        item.Data == parcela.Data.Date && item.Valor == parcela.Valor &&
        MesmaParcela(item.Parcela, parcela.Parcela) &&
        (!DataFoiAjustada(parcela) || DescricaoRelacionada(item.Descricao, parcela.Descricao));

    private static List<LancamentoConferencia> EncontrarIniciosDaSerie(
        ParcelaPluggyConferencia parcela, IReadOnlyList<LancamentoConferencia> lancamentos)
    {
        if (!TryParseParcela(parcela.Parcela, out var numero, out var total) || numero <= 1)
            return new List<LancamentoConferencia>();

        return lancamentos.Where(item => item.IdLancamento.Length > 0 &&
            MesmaParcela(item.Parcela, $"1/{total}") &&
            Igual(item.Pessoa, parcela.Pessoa) && Igual(item.Cartao, parcela.Cartao) &&
            Math.Abs(item.Valor - parcela.Valor) <= 0.10m &&
            Math.Abs((item.Data.AddMonths(numero - 1) - parcela.Data.Date).Days) <= 1).ToList();
    }

    private static bool SerieCompleta(LancamentoConferencia inicio, ParcelaPluggyConferencia parcela,
        IReadOnlyList<LancamentoConferencia> lancamentos)
    {
        if (!TryParseParcela(parcela.Parcela, out _, out var total)) return false;
        var serie = lancamentos.Where(item => item.IdLancamento == inicio.IdLancamento &&
            Igual(item.Pessoa, parcela.Pessoa) && Igual(item.Cartao, parcela.Cartao) &&
            item.Valor == parcela.Valor).ToList();
        return Enumerable.Range(1, total).All(numero => serie.Count(item =>
            MesmaParcela(item.Parcela, $"{numero}/{total}") &&
            item.Data.Year == inicio.Data.AddMonths(numero - 1).Year &&
            item.Data.Month == inicio.Data.AddMonths(numero - 1).Month) == 1);
    }

    private static bool TryParseParcela(string parcela, out int numero, out int total)
    {
        numero = total = 0;
        var partes = parcela.Split('/');
        return partes.Length == 2 && int.TryParse(partes[0], out numero) &&
            int.TryParse(partes[1], out total) && total > 1 && numero >= 1 && numero <= total;
    }

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
