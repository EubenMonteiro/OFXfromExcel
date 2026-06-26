using QifFromExcel;

// ─── CONFIGURATION ────────────────────────────────────────────────────────────
// Full path to the Excel workbook containing the "Money" table.
const string ExcelFilePath =
    @"C:\Users\euben\OneDrive - Euben Silveira Monteiro Junior ME\Pessoal novo\Money\Recibos para o Money.xlsx";

// Folder where QIF files are written. File name: [Card]yyyyMMdd-HHmm.qif
const string OutputFolder =
    @"C:\Users\euben\OneDrive - Euben Silveira Monteiro Junior ME\Pessoal novo\Money\Gerando OFX";
// ──────────────────────────────────────────────────────────────────────────────

try
{
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

    // 2. Load unexported transactions for that card
    var transactions = reader.GetUnexported(selectedCard);
    if (transactions.Count == 0)
    {
        Console.WriteLine($"No unexported transactions found for '{selectedCard}'.");
        return;
    }

    Console.WriteLine($"\nFound {transactions.Count} unexported transaction(s) for '{selectedCard}'.");

    // 3. Generate QIF content (exchange rate prompts happen here for non-BRL rows)
    string qifContent = QifGenerator.Generate(transactions, rateCache);

    // 4. Write QIF file
    Directory.CreateDirectory(OutputFolder);
    var exportedAt = DateTime.Now;
    string safeCard = string.Concat(selectedCard.Split(Path.GetInvalidFileNameChars()));
    string fileName = $"{safeCard}{exportedAt:yyyyMMdd-HHmm}.qif";
    string outputPath = Path.Combine(OutputFolder, fileName);

    File.WriteAllText(outputPath, qifContent, System.Text.Encoding.UTF8);
    Console.WriteLine($"\nQIF file written: {outputPath}");

    // 5. Mark rows as exported in Excel
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
