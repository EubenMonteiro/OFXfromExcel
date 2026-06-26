using System.Text;
using QifFromExcel.Models;

namespace QifFromExcel;

public static class QifGenerator
{
    public static string Generate(IReadOnlyList<Transaction> transactions, ExchangeRateCache rateCache)
    {
        var sb = new StringBuilder();
        sb.AppendLine("!Type:Bank");

        foreach (var tx in transactions)
        {
            decimal brlAmount = tx.Currency == "BRL"
                ? tx.Amount
                : tx.Amount * rateCache.GetRate(tx.Currency);

            // All amounts are debits (positive in Excel = money spent = negative in QIF)
            decimal qifAmount = -brlAmount;

            string memo = BuildMemo(tx);

            sb.AppendLine($"D{tx.Date:dd/MM'yyyy}");
            sb.AppendLine($"T{qifAmount.ToString("F2", System.Globalization.CultureInfo.InvariantCulture)}");
            sb.AppendLine($"P{tx.Payee}");
            if (!string.IsNullOrEmpty(memo))
                sb.AppendLine($"M{memo}");
            sb.AppendLine(BuildCategory(tx));
            sb.AppendLine("^");
        }

        return sb.ToString();
    }

    private static string BuildMemo(Transaction tx)
    {
        var parts = new List<string>();

        if (tx.Currency != "BRL")
        {
            // Prepend original foreign amount so it's visible in Money's memo field.
            // Format matches the Brazilian locale convention: "EUR 5,10"
            string foreignAmount = tx.Amount.ToString("N2", new System.Globalization.CultureInfo("pt-BR"));
            parts.Add($"{tx.Currency} {foreignAmount}");
        }

        if (!string.IsNullOrEmpty(tx.Memo))
            parts.Add(tx.Memo);

        return string.Join(" - ", parts);
    }

    private static string BuildCategory(Transaction tx)
    {
        // QIF L field: "Category" or "Category:Subcategory"
        if (!string.IsNullOrEmpty(tx.SubCategory))
            return $"L{tx.Category}:{tx.SubCategory}";
        if (!string.IsNullOrEmpty(tx.Category))
            return $"L{tx.Category}";
        return "";
    }
}
