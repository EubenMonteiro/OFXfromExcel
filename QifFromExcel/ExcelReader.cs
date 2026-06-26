using ClosedXML.Excel;
using QifFromExcel.Models;

namespace QifFromExcel;

public class ExcelReader : IDisposable
{
    private readonly XLWorkbook _workbook;
    private readonly IXLTable _table;

    // Column names in the Excel table — adjust here if the sheet changes.
    private const string ColDate       = "Data";
    private const string ColPayee      = "Estabelecimento";
    private const string ColAmount     = "Valor";
    private const string ColCurrency   = "Moeda";
    private const string ColCard       = "Cartão";
    private const string ColCategory   = "Categoria";
    private const string ColSubCat     = "Sub-Cat";
    private const string ColMemo       = "Memo";
    private const string ColExported   = "Exportado em";

    private readonly string _filePath;

    public ExcelReader(string filePath)
    {
        _filePath = filePath;
        _workbook = new XLWorkbook(filePath);

        var ws = _workbook.Worksheet("Sheet1");
        _table = ws.Tables.FirstOrDefault(t => t.Name == "Money")
                  ?? throw new InvalidOperationException("Table 'Money' not found in Sheet1.");
    }

    public IReadOnlyList<string> GetCards()
    {
        return _table.DataRange.Rows()
            .Select(r => r.Field(ColCard.Trim()).GetString().Trim())
            .Where(c => !string.IsNullOrEmpty(c))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(c => c)
            .ToList();
    }

    public IReadOnlyList<Transaction> GetUnexported(string card)
    {
        var result = new List<Transaction>();
        int rowNum = 2; // data starts at row 2 (row 1 = headers)

        foreach (var row in _table.DataRange.Rows())
        {
            var exported = row.Field(ColExported.Trim()).GetDateTime();
            if (exported == default) // null / empty = not yet exported
            {
                var cardVal = row.Field(ColCard.Trim()).GetString().Trim();
                if (string.Equals(cardVal, card, StringComparison.OrdinalIgnoreCase))
                {
                    result.Add(new Transaction
                    {
                        RowNumber   = rowNum,
                        Date        = row.Field(ColDate).GetDateTime(),
                        Payee       = row.Field(ColPayee).GetString().Trim(),
                        Amount      = (decimal)row.Field(ColAmount).GetDouble(),
                        Currency    = row.Field(ColCurrency.Trim()).GetString().Trim().ToUpperInvariant(),
                        Card        = cardVal,
                        Category    = row.Field(ColCategory.Trim()).GetString().Trim(),
                        SubCategory = row.Field(ColSubCat).GetString().Trim(),
                        Memo        = row.Field(ColMemo).GetString().Trim(),
                    });
                }
            }
            rowNum++;
        }

        return result;
    }

    public void MarkExported(IEnumerable<Transaction> transactions, DateTime exportedAt)
    {
        var ws = _workbook.Worksheet("Sheet1");

        foreach (var tx in transactions)
        {
            // Row index in the worksheet: header is row 1, data starts at row 2
            var cell = ws.Cell(tx.RowNumber, _table.HeadersRow().Field(ColExported).WorksheetColumn().ColumnNumber());
            cell.Value = exportedAt;
            cell.Style.DateFormat.Format = "dd/MM/yyyy HH:mm";
        }

        _workbook.Save();
    }

    public void Dispose() => _workbook.Dispose();
}
