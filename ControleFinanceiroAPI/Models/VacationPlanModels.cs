namespace ControleFinanceiroAPI.Models;

public sealed class VacationPlanRequest
{
    public DateTime StartDate { get; set; }
    public DateTime EndDate { get; set; }
    public decimal Budget { get; set; }
    public bool UseGuardado { get; set; }
    public decimal? ExpectedLostIncome { get; set; }
    public decimal? ExpectedLowestPeriodBalanceAfter { get; set; }
    public List<VacationParticipant> Participants { get; set; } = new();
    public List<VacationGuardadoTransfer> GuardadoTransfers { get; set; } = new();
    public List<VacationGuardadoTransfer> ExpectedGuardadoTransfers { get; set; } = new();
}

public sealed class VacationGuardadoTransfer
{
    public string Id { get; set; } = string.Empty;
    public int RowNumber { get; set; }
    public string Person { get; set; } = string.Empty;
    public string Period { get; set; } = string.Empty;
    public decimal VacationAmount { get; set; }
    public decimal IncomeAmount { get; set; }
}

public sealed class VacationGuardadoOption
{
    public string Id { get; set; } = string.Empty;
    public int RowNumber { get; set; }
    public string Person { get; set; } = string.Empty;
    public string Period { get; set; } = string.Empty;
    public decimal Available { get; set; }
}

public sealed class VacationParticipant
{
    public string Person { get; set; } = string.Empty;
    public decimal Percent { get; set; }
}

public sealed class VacationPersonPreview
{
    public string Person { get; set; } = string.Empty;
    public decimal Percent { get; set; }
    public decimal BudgetShare { get; set; }
    public decimal HourlyRate { get; set; }
    public int UnpaidHours { get; set; }
    public decimal LostIncome { get; set; }
    public decimal MonthlySaving { get; set; }
    public decimal GuardadoAllocated { get; set; }
    public decimal IncomeGuardadoAllocated { get; set; }
    public decimal IncomeMonthlySaving { get; set; }
    public decimal IncomeReserveFunded { get; set; }
    public decimal TripReserveShortfall { get; set; }
    public decimal IncomeReserveShortfall { get; set; }
    public List<string> ClosedPeriods { get; set; } = new();
    public decimal BalanceBefore { get; set; }
    public decimal BalanceAfter { get; set; }
    public decimal LowestPeriodBalanceAfter { get; set; }
    public string TightestPeriod { get; set; } = string.Empty;
    public Dictionary<string, decimal> LossByPeriod { get; set; } = new();
    public Dictionary<string, decimal> IncomeCoverageByPeriod { get; set; } = new();
}

public sealed class VacationPeriodClosingPreview
{
    public bool HasPlannedVacation { get; set; }
    public bool IsClosed { get; set; }
    public bool CanClose { get; set; }
    public decimal CurrentBalance { get; set; }
    public decimal ReductionNeeded { get; set; }
    public decimal AvailableGuardado { get; set; }
    public decimal ProjectedBalance { get; set; }
    public decimal TripReserveShortfall { get; set; }
    public decimal IncomeReserveShortfall { get; set; }
    public string? Message { get; set; }
}

public sealed class VacationPeriodClosingRequest
{
    public string Person { get; set; } = string.Empty;
    public string Period { get; set; } = string.Empty;
    public decimal ExpectedBalance { get; set; }
    public decimal ExpectedReduction { get; set; }
}

public sealed class VacationOptionPreview
{
    public DateTime StartDate { get; set; }
    public DateTime EndDate { get; set; }
    public int BusinessDays { get; set; }
    public decimal Budget { get; set; }
    public decimal LostIncome { get; set; }
    public decimal BalanceBefore { get; set; }
    public decimal BalanceAfter { get; set; }
    public decimal LowestPeriodBalanceAfter { get; set; }
    public string TightestPeriod { get; set; } = string.Empty;
    public int SavingPeriods { get; set; }
    public bool HasEstimatedSalary { get; set; }
    public List<VacationPersonPreview> People { get; set; } = new();
}

public sealed class VacationPlanPreview
{
    public VacationOptionPreview Selected { get; set; } = new();
    public VacationOptionPreview Recommended { get; set; } = new();
    public List<VacationOptionPreview> Alternatives { get; set; } = new();
    public string? Warning { get; set; }
    public List<VacationGuardadoOption> AvailableGuardado { get; set; } = new();
    public List<VacationGuardadoTransfer> GuardadoTransfers { get; set; } = new();
}
