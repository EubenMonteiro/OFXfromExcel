using System.Text.Json;
using System.Text.Json.Serialization;

namespace QifFromExcel;

/// <summary>
/// Caches exchange rates keyed by currency + 60-day transaction window.
/// A window starts at the earliest transaction date in a group and spans 60 days.
/// The user is prompted once per window; the rate is stored with its exact date range.
/// </summary>
public class ExchangeRateCache
{
    private const int WindowDays = 60;

    private readonly string _cacheFile;
    private List<CachedRate> _rates;

    public ExchangeRateCache(string excelFilePath)
    {
        _cacheFile = Path.ChangeExtension(excelFilePath, ".rates.json");
        _rates = Load();
    }

    /// <summary>
    /// Given a set of transaction dates for one currency, groups them into 60-day windows,
    /// retrieves or prompts for a rate per window, and returns a lookup: date → BRL rate.
    /// </summary>
    public Dictionary<DateTime, decimal> BuildRateLookup(string currency, IEnumerable<DateTime> dates)
    {
        currency = currency.ToUpperInvariant();
        var lookup = new Dictionary<DateTime, decimal>();

        // Sort dates and assign each to a 60-day window anchored at the first date in the group
        var sorted = dates.OrderBy(d => d).ToList();
        if (sorted.Count == 0) return lookup;

        var windows = new List<(DateTime start, DateTime end, List<DateTime> txDates)>();
        DateTime? windowStart = null;

        foreach (var date in sorted)
        {
            if (windowStart == null || (date - windowStart.Value).TotalDays > WindowDays)
            {
                windowStart = date;
                windows.Add((date, date.AddDays(WindowDays), new List<DateTime>()));
            }
            windows[^1].txDates.Add(date);
        }

        // For each window, get or prompt for the rate
        foreach (var (start, end, txDates) in windows)
        {
            decimal rate = GetOrPrompt(currency, start, end);
            foreach (var d in txDates)
                lookup[d] = rate;
        }

        return lookup;
    }

    private decimal GetOrPrompt(string currency, DateTime windowStart, DateTime windowEnd)
    {
        var cached = _rates.FirstOrDefault(r =>
            r.Currency.Equals(currency, StringComparison.OrdinalIgnoreCase) &&
            r.WindowStart == windowStart &&
            r.WindowEnd == windowEnd);

        if (cached != null)
        {
            Console.WriteLine($"Using cached {currency}/BRL rate {cached.Rate} for window {windowStart:dd/MM/yyyy}–{windowEnd:dd/MM/yyyy}.");
            return cached.Rate;
        }

        Console.WriteLine($"\nTransactions in {currency} span {windowStart:dd/MM/yyyy} to {windowEnd:dd/MM/yyyy}.");
        decimal rate = 0;
        while (rate <= 0)
        {
            Console.Write($"Enter {currency}/BRL exchange rate for this period (e.g. 6.25): ");
            var input = Console.ReadLine()?.Replace(',', '.') ?? "";
            if (!decimal.TryParse(input, System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out rate) || rate <= 0)
            {
                Console.WriteLine("Invalid rate. Please enter a positive number.");
                rate = 0;
            }
        }

        _rates.Add(new CachedRate
        {
            Currency    = currency,
            WindowStart = windowStart,
            WindowEnd   = windowEnd,
            Rate        = rate
        });
        Save();
        return rate;
    }

    private List<CachedRate> Load()
    {
        if (!File.Exists(_cacheFile)) return new();
        try
        {
            var json = File.ReadAllText(_cacheFile);
            return JsonSerializer.Deserialize<List<CachedRate>>(json) ?? new();
        }
        catch { return new(); }
    }

    private void Save()
    {
        var json = JsonSerializer.Serialize(_rates, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(_cacheFile, json);
    }

    private class CachedRate
    {
        [JsonPropertyName("currency")]    public string Currency    { get; set; } = "";
        [JsonPropertyName("windowStart")] public DateTime WindowStart { get; set; }
        [JsonPropertyName("windowEnd")]   public DateTime WindowEnd   { get; set; }
        [JsonPropertyName("rate")]        public decimal Rate         { get; set; }
    }
}
