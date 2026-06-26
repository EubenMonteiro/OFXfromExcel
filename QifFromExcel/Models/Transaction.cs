namespace QifFromExcel.Models;

public class Transaction
{
    public int RowNumber { get; init; }
    public DateTime Date { get; init; }
    public string Payee { get; init; } = "";
    public decimal Amount { get; init; }      // original amount in the row's currency
    public string Currency { get; init; } = "BRL";
    public string Card { get; init; } = "";
    public string Category { get; init; } = "";
    public string SubCategory { get; init; } = "";
    public string Memo { get; init; } = "";
    public bool IsExported => ExportedAt.HasValue;
    public DateTime? ExportedAt { get; set; }
}
