using System.Text;
using QifFromExcel.Models;

namespace QifFromExcel;

public static class QifGenerator
{
    public static string Generate(
        IReadOnlyList<Transaction> transactions,
        ExchangeRateCache rateCache,
        string accountCurrency)
    {
        bool isBrlAccount = accountCurrency == "BRL";

        // For BRL accounts: build rate lookups per foreign currency (date → rate)
        Dictionary<string, Dictionary<DateTime, decimal>> rateLookups = new();
        if (isBrlAccount)
        {
            rateLookups = transactions
                .Where(t => t.Currency != "BRL")
                .GroupBy(t => t.Currency)
                .ToDictionary(
                    g => g.Key,
                    g => rateCache.BuildRateLookup(g.Key, g.Select(t => t.Date))
                );
        }

        var sb = new StringBuilder();
        sb.AppendLine("!Type:Bank");

        foreach (var tx in transactions)
        {
            decimal qifAmount;
            string memo;

            if (isBrlAccount)
            {
                decimal brlAmount = tx.Currency == "BRL"
                    ? tx.Amount
                    : tx.Amount * rateLookups[tx.Currency][tx.Date];

                qifAmount = -brlAmount; // positive in Excel = debit on card = negative in QIF
                memo = BuildMemo(tx, prependForeignAmount: tx.Currency != "BRL");
            }
            else
            {
                // Foreign currency account: amounts are already in the account currency, no conversion
                qifAmount = -tx.Amount;
                memo = BuildMemo(tx, prependForeignAmount: false);
            }

            sb.AppendLine($"D{tx.Date.ToString(@"dd/MM\'yyyy")}");
            sb.AppendLine($"T{qifAmount.ToString("F2", System.Globalization.CultureInfo.InvariantCulture)}");
            sb.AppendLine($"P{tx.Payee}");
            if (!string.IsNullOrEmpty(memo))
                sb.AppendLine($"M{memo}");
            string category = BuildCategory(tx);
            if (!string.IsNullOrEmpty(category))
                sb.AppendLine(category);
            sb.AppendLine("^");
        }

        return sb.ToString();
    }

    private static string BuildMemo(Transaction tx, bool prependForeignAmount)
    {
        var parts = new List<string>();

        if (prependForeignAmount)
        {
            string foreignAmount = tx.Amount.ToString("N2", new System.Globalization.CultureInfo("pt-BR"));
            parts.Add($"{tx.Currency} {foreignAmount}");
        }

        if (!string.IsNullOrEmpty(tx.Memo))
            parts.Add(tx.Memo);

        return string.Join(" - ", parts);
    }

    private static string BuildCategory(Transaction tx)
    {
        if (!string.IsNullOrEmpty(tx.Category) && !string.IsNullOrEmpty(tx.SubCategory))
            return $"L{tx.Category}:{tx.SubCategory}";
        if (!string.IsNullOrEmpty(tx.Category))
            return $"L{tx.Category}";
        return "";
    }
}
