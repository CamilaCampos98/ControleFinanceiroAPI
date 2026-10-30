using ControleFinanceiroAPI.Models;
using Google.Apis.Sheets.v4.Data;
using System.Globalization;
using System.Text;

public partial class GoogleSheetsService
{
    public PartySnapshot AddPartyGuest(PartyGuestInput input)
    {
        lock (_lockId)
        {
            var snapshot = RequireParty();
            var guest = BuildGuest(input, snapshot, null);
            var tabs = EnsurePartyTabs();
            AppendPartyRows(tabs, (PartyGuestsSheet, [PartyGuestRow(guest)]));
            return ReadParty();
        }
    }

    public PartySnapshot UpdatePartyGuest(string id, PartyGuestInput input)
    {
        lock (_lockId)
        {
            var snapshot = RequireParty();
            var current = snapshot.Guests.FirstOrDefault(guest => guest.Id == id)
                ?? throw new ArgumentException("Convidado não encontrado.");
            var guest = BuildGuest(input, snapshot, current);
            var tabs = EnsurePartyTabs();
            WritePartyRow(PartyGuestsSheet, FindPartyRow(PartyGuestsSheet, tabs, id) + 1, PartyGuestRow(guest));
            return ReadParty();
        }
    }

    public PartySnapshot DeletePartyGuest(string id)
    {
        lock (_lockId)
        {
            var snapshot = RequireParty();
            if (!snapshot.Guests.Any(guest => guest.Id == id)) throw new ArgumentException("Convidado não encontrado.");
            var tabs = EnsurePartyTabs();
            DeletePartyRow(tabs[PartyGuestsSheet], FindPartyRow(PartyGuestsSheet, tabs, id));
            return ReadParty();
        }
    }

    public PartySnapshot AddPartyHouse(PartyHouseInput input)
    {
        lock (_lockId)
        {
            var snapshot = RequireParty();
            var name = ValidateHouseName(input.Name);
            var selected = ValidateHouseGuestIds(snapshot, input.GuestIds);
            var now = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture);
            var house = new PartyHouse(Guid.NewGuid().ToString(), snapshot.Event!.Id, name, now, now);
            var tabs = EnsurePartyTabs();
            var rows = PartyRows(PartyGuestsSheet, tabs);
            var requests = new List<Request> { PartyAppendRequest(tabs[PartyHousesSheet], PartyHouseRow(house)) };
            requests.AddRange(HouseAssignmentUpdates(tabs[PartyGuestsSheet], rows, house.Id, selected));
            _service.Spreadsheets.BatchUpdate(new BatchUpdateSpreadsheetRequest { Requests = requests }, SpreadsheetId).Execute();
            return ReadParty();
        }
    }

    public PartySnapshot UpdatePartyHouse(string id, PartyHouseInput input)
    {
        lock (_lockId)
        {
            var snapshot = RequireParty();
            var current = snapshot.Houses.FirstOrDefault(house => house.Id == id)
                ?? throw new ArgumentException("Casa não encontrada.");
            var name = ValidateHouseName(input.Name);
            var selected = ValidateHouseGuestIds(snapshot, input.GuestIds ?? snapshot.Guests
                .Where(guest => guest.HouseId == id).Select(guest => guest.Id).ToList());
            var now = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture);
            var tabs = EnsurePartyTabs();
            var rows = PartyRows(PartyGuestsSheet, tabs);
            var requests = PartyRowUpdates(tabs[PartyHousesSheet], FindPartyRow(PartyHousesSheet, tabs, id),
                PartyHouseRow(current with { Name = name, UpdatedAt = now }));
            requests.AddRange(HouseAssignmentUpdates(tabs[PartyGuestsSheet], rows, id, selected));
            _service.Spreadsheets.BatchUpdate(new BatchUpdateSpreadsheetRequest { Requests = requests }, SpreadsheetId).Execute();
            return ReadParty();
        }
    }

    public PartySnapshot DeletePartyHouse(string id)
    {
        lock (_lockId)
        {
            var snapshot = RequireParty();
            if (!snapshot.Houses.Any(house => house.Id == id)) throw new ArgumentException("Casa não encontrada.");
            var tabs = EnsurePartyTabs();
            var rows = PartyRows(PartyGuestsSheet, tabs);
            var requests = HouseAssignmentUpdates(tabs[PartyGuestsSheet], rows, id, new HashSet<string>());
            var rowIndex = FindPartyRow(PartyHousesSheet, tabs, id);
            requests.Add(new Request { DeleteDimension = new DeleteDimensionRequest
            {
                Range = new DimensionRange { SheetId = tabs[PartyHousesSheet], Dimension = "ROWS",
                    StartIndex = rowIndex, EndIndex = rowIndex + 1 }
            } });
            _service.Spreadsheets.BatchUpdate(new BatchUpdateSpreadsheetRequest { Requests = requests }, SpreadsheetId).Execute();
            return ReadParty();
        }
    }

    private static PartyGuest BuildGuest(PartyGuestInput input, PartySnapshot snapshot, PartyGuest? current)
    {
        var name = input.Name?.Trim() ?? "";
        if (name.Length is < 1 or > 120) throw new ArgumentException("Informe o nome do convidado (até 120 caracteres).");
        if (current == null || NormalizePartyName(current.Name) != NormalizePartyName(name))
        {
            if (snapshot.Guests.Any(guest => guest.Id != current?.Id && NormalizePartyName(guest.Name) == NormalizePartyName(name)))
                throw new ArgumentException($"{name} já está cadastrado na lista.");
        }
        var gender = input.IsChild ? input.Gender?.Trim() ?? "" : "";
        if (input.IsChild && gender is not ("Menino" or "Menina"))
            throw new ArgumentException("Informe se a criança é menino ou menina.");
        var houseId = input.HouseId?.Trim() ?? "";
        if (houseId.Length > 0 && !snapshot.Houses.Any(house => house.Id == houseId))
            throw new ArgumentException("Casa não encontrada.");
        var notes = input.Notes?.Trim() ?? "";
        if (notes.Length > 500) throw new ArgumentException("A observação deve ter no máximo 500 caracteres.");
        var now = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture);
        return new PartyGuest(current?.Id ?? Guid.NewGuid().ToString(), snapshot.Event!.Id, name,
            input.IsChild, gender, notes, current?.CreatedAt ?? now, now, houseId);
    }

    private static string ValidateHouseName(string? value)
    {
        var name = value?.Trim() ?? "";
        if (name.Length is < 1 or > 120) throw new ArgumentException("Informe o nome da casa (até 120 caracteres).");
        return name;
    }

    private static HashSet<string> ValidateHouseGuestIds(PartySnapshot snapshot, List<string>? guestIds)
    {
        var selected = (guestIds ?? new List<string>()).ToHashSet(StringComparer.Ordinal);
        if (selected.Any(id => !snapshot.Guests.Any(guest => guest.Id == id)))
            throw new ArgumentException("Há convidado não encontrado na seleção da casa.");
        return selected;
    }

    private static string NormalizePartyName(string value)
    {
        var decomposed = value.Trim().Normalize(NormalizationForm.FormD);
        var builder = new StringBuilder();
        var wasSpace = false;
        foreach (var character in decomposed)
        {
            if (CharUnicodeInfo.GetUnicodeCategory(character) == UnicodeCategory.NonSpacingMark) continue;
            if (char.IsWhiteSpace(character))
            {
                if (!wasSpace) builder.Append(' ');
                wasSpace = true;
            }
            else { builder.Append(char.ToLowerInvariant(character)); wasSpace = false; }
        }
        return builder.ToString();
    }

    private int FindPartyRow(string sheet, Dictionary<string, int> tabs, string id)
    {
        var index = PartyRows(sheet, tabs).FindIndex(row => Cell(row, 0) == id);
        return index > 0 ? index : throw new ArgumentException("Registro não encontrado na planilha.");
    }

    private void DeletePartyRow(int sheetId, int rowIndex) => _service.Spreadsheets.BatchUpdate(
        new BatchUpdateSpreadsheetRequest { Requests = [new Request { DeleteDimension = new DeleteDimensionRequest
        {
            Range = new DimensionRange { SheetId = sheetId, Dimension = "ROWS", StartIndex = rowIndex,
                EndIndex = rowIndex + 1 }
        } }] }, SpreadsheetId).Execute();

    private static List<Request> HouseAssignmentUpdates(int sheetId, List<IList<object>> rows,
        string houseId, HashSet<string> selected)
    {
        var requests = new List<Request>();
        for (var index = 1; index < rows.Count; index++)
        {
            var guestId = Cell(rows[index], 0);
            var previous = Cell(rows[index], 8);
            var next = selected.Contains(guestId) ? houseId : previous == houseId ? "" : previous;
            if (next == previous) continue;
            requests.Add(PartyCellUpdate(sheetId, index, 8, next));
            requests.Add(PartyCellUpdate(sheetId, index, 7, DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture)));
        }
        return requests;
    }

    private static Request PartyAppendRequest(int sheetId, IList<object> row) => new()
    {
        AppendCells = new AppendCellsRequest { SheetId = sheetId, Fields = "userEnteredValue", Rows =
            [new RowData { Values = row.Select(value => new CellData
            { UserEnteredValue = new ExtendedValue { StringValue = value?.ToString() ?? "" } }).ToList() }] }
    };

    private static List<Request> PartyRowUpdates(int sheetId, int rowIndex, IList<object> row) =>
        row.Select((value, column) => PartyCellUpdate(sheetId, rowIndex, column, value?.ToString() ?? "")).ToList();

    private static IList<object> PartyGuestRow(PartyGuest guest) =>
        [guest.Id, guest.EventId, guest.Name, guest.IsChild ? "Sim" : "Não", guest.Gender,
            guest.Notes, guest.CreatedAt, guest.UpdatedAt, guest.HouseId];

    private static IList<object> PartyHouseRow(PartyHouse house) =>
        [house.Id, house.EventId, house.Name, house.CreatedAt, house.UpdatedAt];
}
