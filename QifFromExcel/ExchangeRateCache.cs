using System.Text.Json;

namespace QifFromExcel;

/// <summary>
/// Persists exchange rates to a JSON file next to the Excel workbook.
/// A cached rate is reused for 60 days before the user is re-prompted.
/// </summary>
public class ExchangeRateCache
{
    private const int CacheWindowDays = 60;

    private readonly string _cacheFile;
    private Dictionary<string, CachedRate> _rates;

    public ExchangeRateCache(string excelFilePath)
    {
        _cacheFile = Path.ChangeExtension(excelFilePath, ".rates.json");
        _rates = Load();
    }

    public decimal GetRate(string currency)
    {
        currency = currency.ToUpperInvariant();

        if (_rates.TryGetValue(currency, out var cached) &&
            (DateTime.Today - cached.AsOf).TotalDays <= CacheWindowDays)
        {
            Console.WriteLine($"Using cached {currency}/BRL rate: {cached.Rate} (set on {cached.AsOf:dd/MM/yyyy}, valid for {CacheWindowDays} days).");
            return cached.Rate;
        }

        return PromptAndSave(currency);
    }

    private decimal PromptAndSave(string currency)
    {
        decimal rate = 0;
        while (rate <= 0)
        {
            Console.Write($"Enter exchange rate for {currency} → BRL (e.g. 6.25): ");
            var input = Console.ReadLine()?.Replace(',', '.') ?? "";
            if (!decimal.TryParse(input, System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out rate) || rate <= 0)
            {
                Console.WriteLine("Invalid rate. Please enter a positive number.");
                rate = 0;
            }
        }

        _rates[currency] = new CachedRate { Rate = rate, AsOf = DateTime.Today };
        Save();
        return rate;
    }

    private Dictionary<string, CachedRate> Load()
    {
        if (!File.Exists(_cacheFile)) return new();
        try
        {
            var json = File.ReadAllText(_cacheFile);
            return JsonSerializer.Deserialize<Dictionary<string, CachedRate>>(json) ?? new();
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
        public decimal Rate { get; set; }
        public DateTime AsOf { get; set; }
    }
}
