using ControleFinanceiroAPI.Models;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System.Globalization;

public partial class GoogleSheetsService
{
    private const string PartyEventsSheet = "Festas";
    private const string PartyItemsSheet = "FestaItens";
    private const string PartyInstallmentsSheet = "FestaParcelas";
    private const string PartyPaymentsSheet = "FestaPagamentos";
    private const string LiviaPartyId = "livia-1-ano-2027";

    private static readonly Dictionary<string, string[]> PartyHeaders = new()
    {
        [PartyEventsSheet] = ["Id", "Nome", "Data"],
        [PartyItemsSheet] = ["Id", "FestaId", "Categoria", "Item", "Valor"],
        [PartyInstallmentsSheet] = ["Id", "FestaId", "Categoria", "ItemId", "Vencimento", "Valor"],
        [PartyPaymentsSheet] = ["Id", "FestaId", "Categoria", "ItemId", "ParcelaId", "Descricao", "DataPagamento", "Valor"]
    };

    public PartySnapshot ReadParty()
    {
        lock (_lockId)
        {
            var tabs = PartyTabIds();
            if (!tabs.ContainsKey(PartyEventsSheet)) return new PartySnapshot();
            var data = ReadPartyBatch(tabs);
            var events = data[PartyEventsSheet];
            var party = events.Skip(1).Select(row => new PartyEvent(Cell(row, 0), Cell(row, 1), ParsePartyDate(Cell(row, 2))))
                .FirstOrDefault(item => item.Id == LiviaPartyId);
            if (party == null) return new PartySnapshot();
            var items = data[PartyItemsSheet].Skip(1)
                .Where(row => Cell(row, 1) == party.Id)
                .Select(row => new PartyItem(Cell(row, 0), Cell(row, 1), Cell(row, 2), Cell(row, 3),
                    string.IsNullOrWhiteSpace(Cell(row, 4)) ? null : ParsePartyMoney(Cell(row, 4)))).ToList();
            var installments = data[PartyInstallmentsSheet].Skip(1)
                .Where(row => Cell(row, 1) == party.Id)
                .Select(row => new PartyInstallment(Cell(row, 0), Cell(row, 1), Cell(row, 2), EmptyToNull(Cell(row, 3)),
                    ParsePartyDate(Cell(row, 4)), ParsePartyMoney(Cell(row, 5)))).ToList();
            var payments = data[PartyPaymentsSheet].Skip(1)
                .Where(row => Cell(row, 1) == party.Id)
                .Select(row => new PartyPayment(Cell(row, 0), Cell(row, 1), Cell(row, 2), EmptyToNull(Cell(row, 3)),
                    EmptyToNull(Cell(row, 4)), Cell(row, 5), ParsePartyDate(Cell(row, 6)), ParsePartyMoney(Cell(row, 7)))).ToList();
            var snapshot = new PartySnapshot { Event = party, Items = items, Installments = installments, Payments = payments };
            var labeledPayments = payments.Select(payment =>
            {
                var installment = installments.FirstOrDefault(item => item.Id == payment.InstallmentId);
                return installment == null || payment.Description != "Pagamento de parcela"
                    ? payment
                    : payment with { Description = ControleFinanceiroAPI.Services.PartyPaymentNaming.Describe(
                        snapshot, installment, payment.Amount) };
            }).ToList();
            return new PartySnapshot { Event = party, Items = items, Installments = installments, Payments = labeledPayments };
        }
    }

    public PartySnapshot ImportLiviaParty()
    {
        lock (_lockId)
        {
            var existing = ReadParty();
            if (existing.Event != null) return existing;
            var tabs = EnsurePartyTabs();
            var party = new PartyEvent(LiviaPartyId, "1º aniversário da Lívia", new DateTime(2027, 3, 27));
            var items = new[]
            {
                new PartyItem("livia-espaco", party.Id, "Salão", "Espaço", 700),
                new PartyItem("livia-toalhas", party.Id, "Salão", "Toalhas de mesa", 160),
                new PartyItem("livia-decoracao", party.Id, "Salão", "Decoração", 750),
                new PartyItem("livia-baloes", party.Id, "Salão", "Balões dourados", 60),
                new PartyItem("livia-brinquedos", party.Id, "Salão", "Brinquedos", 300),
                new PartyItem("livia-kit-pintura", party.Id, "Itens necessários", "Kit Pintura", null),
                new PartyItem("livia-saco-bombom", party.Id, "Itens necessários", "Saco de bombom", null),
                new PartyItem("livia-lembrancinhas", party.Id, "Itens necessários", "Lembrancinhas", 597),
                new PartyItem("livia-doces-personalizados", party.Id, "Itens necessários", "Doces personalizados", null)
            };
            // Os pagamentos do salão pertencem à categoria, não a um dos cinco itens isoladamente.
            var payments = new[]
            {
                new PartyPayment("livia-sinal-salao", party.Id, "Salão", null, null, "Sinal salão", new DateTime(2026, 8, 13), 200),
                new PartyPayment("livia-parcela-salao", party.Id, "Salão", null, null, "Parcela 1 salão", new DateTime(2026, 9, 15), 200),
                new PartyPayment("livia-parcela-lembrancinhas", party.Id, "Itens necessários", "livia-lembrancinhas", null,
                    "Parcela 1 lembrancinhas", new DateTime(2026, 9, 15), 200)
            };
            var salonSplit = ControleFinanceiroAPI.Services.PartyScheduleCalculator.Split(1570, 6, 0);
            var installments = Enumerable.Range(0, 6).Select(index =>
                new PartyInstallment($"livia-salao-{index + 1}", party.Id, "Salão", null,
                    new DateTime(2026, 10, 15).AddMonths(index), salonSplit[index])).ToArray();
            AppendPartyRows(tabs,
                (PartyEventsSheet, [PartyEventRow(party)]),
                (PartyItemsSheet, items.Select(PartyItemRow).ToList()),
                (PartyInstallmentsSheet, installments.Select(PartyInstallmentRow).ToList()),
                (PartyPaymentsSheet, payments.Select(PartyPaymentRow).ToList()));
            return ReadParty();
        }
    }

    public PartySnapshot AddPartyItem(PartyItemInput input)
    {
        lock (_lockId)
        {
            var snapshot = RequireParty();
            var category = input.Category?.Trim() ?? "";
            var name = input.Name?.Trim() ?? "";
            if (category.Length is < 2 or > 80 || name.Length is < 2 or > 120 || input.Amount is < 0 ||
                input.Amount is { } amount && Math.Round(amount, 2) != amount)
                throw new ArgumentException("Informe categoria, item e valor válido (até duas casas decimais).");
            var tabs = PartyTabIds();
            var item = new PartyItem(Guid.NewGuid().ToString("N"), snapshot.Event!.Id, category, name, input.Amount);
            AppendPartyRows(tabs, (PartyItemsSheet, [PartyItemRow(item)]));
            return ReadParty();
        }
    }

    public PartySnapshot UpdatePartyItem(string id, PartyItemInput input)
    {
        lock (_lockId)
        {
            var snapshot = RequireParty();
            var existing = snapshot.Items.FirstOrDefault(item => item.Id == id)
                ?? throw new ArgumentException("Item não encontrado.");
            var category = input.Category?.Trim() ?? "";
            var name = input.Name?.Trim() ?? "";
            if (category != existing.Category &&
                (snapshot.Installments.Any(item => item.ItemId == id || item.Category == existing.Category && item.ItemId == null) ||
                 snapshot.Payments.Any(item => item.ItemId == id || item.Category == existing.Category && item.ItemId == null)))
                throw new ArgumentException("A categoria de um item com parcelas ou pagamentos não pode ser alterada.");
            var paid = snapshot.Payments.Where(payment => payment.ItemId == id).Sum(payment => payment.Amount);
            if (category.Length is < 2 or > 80 || name.Length is < 2 or > 120 || input.Amount is < 0 ||
                input.Amount is { } amount && (Math.Round(amount, 2) != amount || amount < paid))
                throw new ArgumentException("O valor não pode ser menor que o já pago; use até duas casas decimais.");
            var categoryCommitted = snapshot.Payments.Where(item => item.Category == existing.Category && item.InstallmentId == null)
                .Sum(item => item.Amount) + snapshot.Installments.Where(item => item.Category == existing.Category)
                .Sum(item => item.Amount);
            var categoryNewTotal = snapshot.Items.Where(item => item.Category == existing.Category && item.Id != id)
                .Sum(item => item.Amount ?? 0) + (input.Amount ?? 0);
            if (category == existing.Category && categoryNewTotal < categoryCommitted)
                throw new ArgumentException("O novo valor deixaria pagamentos ou parcelas maiores que o orçamento desta categoria.");
            var tabs = PartyTabIds();
            var rows = PartyRows(PartyItemsSheet, tabs);
            var rowIndex = rows.FindIndex(row => Cell(row, 0) == id);
            if (rowIndex < 1) throw new ArgumentException("Item não encontrado.");
            WritePartyRow(PartyItemsSheet, rowIndex + 1,
                PartyItemRow(existing with { Category = category, Name = name, Amount = input.Amount }));
            return ReadParty();
        }
    }

    public PartySnapshot SchedulePartyPayments(PartyScheduleInput input)
    {
        lock (_lockId)
        {
            var snapshot = RequireParty();
            var category = input.Category?.Trim() ?? "";
            var targetItems = snapshot.Items.Where(item => item.Category == category &&
                (input.ItemId == null || item.Id == input.ItemId)).ToList();
            if (targetItems.Count == 0 || targetItems.Any(item => !item.Amount.HasValue))
                throw new ArgumentException("Defina o valor dos itens antes de criar parcelas.");
            if (input.ItemId != null && targetItems.Count != 1)
                throw new ArgumentException("Item não encontrado.");
            if (input.ItemId != null && snapshot.Installments.Any(item =>
                    item.Category == category && item.ItemId == null))
                throw new ArgumentException("Esta categoria já tem parcelas conjuntas. Planeje o saldo adicional pela categoria.");
            var postedMonths = input.Months ?? new List<DateTime>();
            var months = postedMonths.Select(month => new DateTime(month.Year, month.Month, 1))
                .Distinct().OrderBy(month => month).ToList();
            var today = ControleFinanceiroAPI.Services.VacationPlanningService.TodayBrazil;
            if (months.Count == 0 || months.Count != postedMonths.Count ||
                months.Any(month => month < new DateTime(today.Year, today.Month, 1) || month > snapshot.Event!.Date))
                throw new ArgumentException("Escolha meses distintos entre o mês atual e a data da festa.");
            if (input.DueDay is < 1 or > 31)
                throw new ArgumentException("Informe um dia previsto entre 1 e 31.");
            if (months.Any(month => new DateTime(month.Year, month.Month,
                    Math.Min(input.DueDay, DateTime.DaysInMonth(month.Year, month.Month))) > snapshot.Event!.Date))
                throw new ArgumentException("Todas as parcelas precisam vencer até a data da festa.");
            var remaining = ControleFinanceiroAPI.Services.PartyBalanceCalculator.Unplanned(
                snapshot, category, input.ItemId);
            if (remaining <= 0) throw new ArgumentException("Todo o valor conhecido já está pago ou planejado.");
            var amounts = ControleFinanceiroAPI.Services.PartyScheduleCalculator.Split(remaining, months.Count,
                input.MinimumInstallment);
            var installments = months.Select((month, index) => new PartyInstallment(Guid.NewGuid().ToString("N"),
                snapshot.Event!.Id, category, input.ItemId,
                new DateTime(month.Year, month.Month, Math.Min(input.DueDay, DateTime.DaysInMonth(month.Year, month.Month))),
                amounts[index])).ToList();
            AppendPartyRows(PartyTabIds(), (PartyInstallmentsSheet, installments.Select(PartyInstallmentRow).ToList()));
            return ReadParty();
        }
    }

    public PartySnapshot PayPartyInstallment(string installmentId, string paymentId, DateTime paidAt, decimal amount)
    {
        lock (_lockId)
        {
            var snapshot = RequireParty();
            if (!Guid.TryParseExact(paymentId, "N", out _))
                throw new ArgumentException("Identificador de pagamento inválido.");
            if (snapshot.Payments.Any(item => item.Id == paymentId)) return snapshot;
            var installment = snapshot.Installments.FirstOrDefault(item => item.Id == installmentId)
                ?? throw new ArgumentException("Parcela não encontrada.");
            var paid = snapshot.Payments.Where(item => item.InstallmentId == installmentId).Sum(item => item.Amount);
            if (paidAt.Date > ControleFinanceiroAPI.Services.VacationPlanningService.TodayBrazil ||
                paidAt.Date < new DateTime(2020, 1, 1) || amount <= 0 ||
                Math.Round(amount, 2) != amount || amount > installment.Amount - paid)
                throw new ArgumentException("Informe uma data válida e valor até o saldo da parcela.");
            var description = ControleFinanceiroAPI.Services.PartyPaymentNaming.Describe(snapshot, installment, amount);
            var payment = new PartyPayment(paymentId, snapshot.Event!.Id, installment.Category,
                installment.ItemId, installment.Id, description, paidAt.Date, amount);
            AppendPartyRows(PartyTabIds(), (PartyPaymentsSheet, [PartyPaymentRow(payment)]));
            return ReadParty();
        }
    }

    private PartySnapshot RequireParty()
    {
        var snapshot = ReadParty();
        return snapshot.Event == null
            ? throw new InvalidOperationException("Importe a festa antes de cadastrar itens e parcelas.")
            : snapshot;
    }

    private Dictionary<string, int> PartyTabIds() => _service.Spreadsheets.Get(SpreadsheetId).Execute().Sheets
        .Where(sheet => PartyHeaders.ContainsKey(sheet.Properties.Title))
        .ToDictionary(sheet => sheet.Properties.Title, sheet => sheet.Properties.SheetId!.Value);

    private Dictionary<string, int> EnsurePartyTabs()
    {
        var tabs = PartyTabIds();
        foreach (var (name, header) in PartyHeaders)
        {
            if (!tabs.ContainsKey(name))
            {
                var created = _service.Spreadsheets.BatchUpdate(new BatchUpdateSpreadsheetRequest
                {
                    Requests = [new Request { AddSheet = new AddSheetRequest { Properties = new SheetProperties { Title = name } } }]
                }, SpreadsheetId).Execute();
                tabs[name] = created.Replies[0].AddSheet.Properties.SheetId!.Value;
            }
            if (PartyRows(name, tabs).Count == 0)
                WritePartyRow(name, 1, header.Select(value => (object)value).ToList());
        }
        return tabs;
    }

    private List<IList<object>> PartyRows(string name, Dictionary<string, int> tabs)
    {
        if (!tabs.ContainsKey(name)) return new();
        var rows = ReadData($"{name}!A1:K")?.ToList() ?? new();
        ValidatePartyHeader(name, rows);
        return rows;
    }

    private Dictionary<string, List<IList<object>>> ReadPartyBatch(Dictionary<string, int> tabs)
    {
        var names = PartyHeaders.Keys.Where(tabs.ContainsKey).ToList();
        var request = _service.Spreadsheets.Values.BatchGet(SpreadsheetId);
        request.Ranges = names.Select(name => $"{name}!A1:K").ToList();
        var values = request.Execute().ValueRanges ?? new List<ValueRange>();
        var result = PartyHeaders.Keys.ToDictionary(name => name, _ => new List<IList<object>>());
        for (var index = 0; index < names.Count; index++)
        {
            var rows = values.ElementAtOrDefault(index)?.Values?.ToList() ?? new List<IList<object>>();
            ValidatePartyHeader(names[index], rows);
            result[names[index]] = rows;
        }
        return result;
    }

    private static void ValidatePartyHeader(string name, List<IList<object>> rows)
    {
        if (rows.Count > 0 && PartyHeaders[name].Where((value, index) => Cell(rows[0], index) != value).Any())
            throw new InvalidOperationException($"A aba {name} já existe com colunas diferentes das esperadas.");
    }

    private void WritePartyRow(string name, int rowNumber, IList<object> row)
    {
        var endColumn = (char)('A' + row.Count - 1);
        var request = _service.Spreadsheets.Values.Update(
            new ValueRange { Values = [row] }, SpreadsheetId, $"{name}!A{rowNumber}:{endColumn}{rowNumber}");
        request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
        request.Execute();
    }

    private void AppendPartyRows(Dictionary<string, int> tabs,
        params (string Name, List<IList<object>> Rows)[] groups)
    {
        var requests = groups.Where(group => group.Rows.Count > 0).Select(group => new Request
        {
            AppendCells = new AppendCellsRequest
            {
                SheetId = tabs[group.Name],
                Rows = group.Rows.Select(row => new RowData
                {
                    Values = row.Select(value => new CellData
                    {
                        UserEnteredValue = new ExtendedValue { StringValue = value?.ToString() ?? "" }
                    }).ToList()
                }).ToList(),
                Fields = "userEnteredValue"
            }
        }).ToList();
        _service.Spreadsheets.BatchUpdate(new BatchUpdateSpreadsheetRequest { Requests = requests }, SpreadsheetId).Execute();
    }

    private static IList<object> PartyEventRow(PartyEvent item) => [item.Id, item.Name, PartyDate(item.Date)];
    private static IList<object> PartyItemRow(PartyItem item) =>
        [item.Id, item.EventId, item.Category, item.Name, PartyMoney(item.Amount)];
    private static IList<object> PartyInstallmentRow(PartyInstallment item) =>
        [item.Id, item.EventId, item.Category, item.ItemId ?? "", PartyDate(item.DueDate), PartyMoney(item.Amount)];
    private static IList<object> PartyPaymentRow(PartyPayment item) =>
        [item.Id, item.EventId, item.Category, item.ItemId ?? "", item.InstallmentId ?? "", item.Description,
            PartyDate(item.PaidAt), PartyMoney(item.Amount)];
    private static string PartyDate(DateTime value) => value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
    private static string PartyMoney(decimal? value) => value?.ToString("0.00", CultureInfo.InvariantCulture) ?? "";
    private static string Cell(IList<object> row, int index) => row.ElementAtOrDefault(index)?.ToString()?.Trim() ?? "";
    private static string? EmptyToNull(string value) => string.IsNullOrWhiteSpace(value) ? null : value;
    private static DateTime ParsePartyDate(string value) => DateTime.ParseExact(value, "yyyy-MM-dd", CultureInfo.InvariantCulture);
    private static decimal ParsePartyMoney(string value) => decimal.Parse(value, CultureInfo.InvariantCulture);
}
