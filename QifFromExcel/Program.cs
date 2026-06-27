using System.Text.RegularExpressions;
using QifFromExcel;
using QifFromExcel.Models;

// ─── CONFIGURATION ────────────────────────────────────────────────────────────
// Full path to the Excel workbook containing the "Money" table.
// TEST: pointing at the test copy in the repo folder.
// PRODUCTION: change to @"C:\Users\euben\OneDrive - Euben Silveira Monteiro Junior ME\Pessoal novo\Money\Recibos para o Money.xlsx"
const string ExcelFilePath =
    @"C:\Users\euben\OneDrive - Euben Silveira Monteiro Junior ME\Pessoal novo\Money\Gerando OFX\ofxfromexcel\Recibos para o Money - cópia teste.xlsx";

// Folder where QIF files are written.
const string OutputFolder =
    @"C:\Users\euben\OneDrive - Euben Silveira Monteiro Junior ME\Pessoal novo\Money\Gerando OFX";
// ──────────────────────────────────────────────────────────────────────────────

try
{
    System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
    Console.OutputEncoding = System.Text.Encoding.UTF8;

    using var reader = new ExcelReader(ExcelFilePath);
    var rateCache = new ExchangeRateCache(ExcelFilePath);

    // 1. Ask which card to export
    var cards = reader.GetCards();
    if (cards.Count == 0)
    {
        Console.WriteLine("No cards found in the Money table. Nothing to export.");
        return;
    }

    Console.WriteLine("Cards available for export:");
    for (int i = 0; i < cards.Count; i++)
        Console.WriteLine($"  {i + 1}. {cards[i]}");

    int choice = 0;
    while (choice < 1 || choice > cards.Count)
    {
        Console.Write($"Choose a card (1–{cards.Count}): ");
        int.TryParse(Console.ReadLine(), out choice);
    }
    string selectedCard = cards[choice - 1];

    // 2. Determine account currency
    // Check if the card name already contains a currency code (e.g. "Avenue USD")
    string? impliedCurrency = DetectCurrencyInName(selectedCard);
    string accountCurrency;

    if (impliedCurrency != null)
    {
        Console.WriteLine($"\nCard name suggests currency: {impliedCurrency}. Using it as the account currency.");
        accountCurrency = impliedCurrency;
    }
    else
    {
        Console.Write("\nAccount base currency (press Enter for BRL, or type EUR / USD / etc.): ");
        var input = Console.ReadLine()?.Trim().ToUpperInvariant() ?? "";
        accountCurrency = string.IsNullOrEmpty(input) ? "BRL" : input;
    }

    // 3. Load unexported transactions
    IReadOnlyList<Models.Transaction> transactions;
    if (accountCurrency == "BRL")
    {
        // BRL account: export all currencies, convert to BRL
        transactions = reader.GetUnexported(selectedCard);
    }
    else
    {
        // Foreign currency account: only export rows in that currency
        var allForCard = reader.GetUnexported(selectedCard);

        // Warn if card name implies a different currency than what the user typed
        if (impliedCurrency == null)
        {
            var otherCurrencies = allForCard
                .Where(t => t.Currency != accountCurrency)
                .Select(t => t.Currency)
                .Distinct()
                .ToList();
            if (otherCurrencies.Count > 0)
                Console.WriteLine($"Warning: {otherCurrencies.Count} transaction(s) in other currencies " +
                                  $"({string.Join(", ", otherCurrencies)}) will be skipped — " +
                                  $"they belong in a BRL export run.");
        }

        transactions = allForCard.Where(t => t.Currency == accountCurrency).ToList();
    }

    if (transactions.Count == 0)
    {
        Console.WriteLine($"No unexported transactions found for '{selectedCard}'" +
                          (accountCurrency != "BRL" ? $" in {accountCurrency}" : "") + ".");
        return;
    }

    Console.WriteLine($"\nFound {transactions.Count} unexported transaction(s) for '{selectedCard}'" +
                      (accountCurrency != "BRL" ? $" in {accountCurrency}" : " (all currencies)") + ".");

    // 4. Generate QIF
    string qifContent = QifGenerator.Generate(transactions, rateCache, accountCurrency);

    // 5. Write QIF file
    Directory.CreateDirectory(OutputFolder);
    var exportedAt = DateTime.Now;
    string safeCard = string.Concat(selectedCard.Split(Path.GetInvalidFileNameChars()));
    string fileName = accountCurrency == "BRL"
        ? $"{safeCard}{exportedAt:yyyyMMdd-HHmm}.qif"
        : $"{safeCard} - {accountCurrency} - {exportedAt:yyyyMMdd-HHmm}.qif";
    string outputPath = Path.Combine(OutputFolder, fileName);

    File.WriteAllText(outputPath, qifContent, System.Text.Encoding.GetEncoding(1252));
    Console.WriteLine($"\nQIF file written: {outputPath}");

    // 6. Mark rows as exported in Excel
    reader.MarkExported(transactions, exportedAt);
    Console.WriteLine("Excel rows marked as exported.");

    Console.WriteLine("\nDone. Press any key to exit.");
    Console.ReadKey();
}
catch (Exception ex)
{
    Console.Error.WriteLine($"Error: {ex.Message}");
    Console.Error.WriteLine(ex.StackTrace);
    Console.WriteLine("\nPress any key to exit.");
    Console.ReadKey();
}

// Looks for a 3-letter uppercase currency code at the end of the card name (e.g. "Avenue USD")
static string? DetectCurrencyInName(string cardName)
{
    var match = Regex.Match(cardName, @"\b([A-Z]{3})\s*$");
    if (!match.Success) return null;
    string candidate = match.Groups[1].Value;
    // Exclude common non-currency suffixes that happen to be 3 letters
    return candidate == "BRL" ? null : candidate;
}
