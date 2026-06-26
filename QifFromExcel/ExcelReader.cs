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

    // Column indices (1-based) resolved once from the header row
    private readonly Dictionary<string, int> _colIndex;

    public ExcelReader(string filePath)
    {
        _filePath = filePath;
        _workbook = new XLWorkbook(filePath);

        var ws = _workbook.Worksheet("Sheet1");
        _table = ws.Tables.FirstOrDefault(t => t.Name == "Money")
                  ?? throw new InvalidOperationException("Table 'Money' not found in Sheet1.");

        // Build a map from header label (trimmed) → 1-based column index within the table data range
        _colIndex = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var headerRow = _table.HeadersRow();
        foreach (var cell in headerRow.Cells())
        {
            var label = cell.GetString().Trim();
            if (!string.IsNullOrEmpty(label))
                _colIndex[label] = cell.WorksheetColumn().ColumnNumber();
        }
    }

    private IXLCell DataCell(IXLRangeRow row, string columnName)
    {
        if (!_colIndex.TryGetValue(columnName.Trim(), out int colNum))
            throw new InvalidOperationException($"Column '{columnName}' not found in table 'Money'.");
        return row.WorksheetRow().Cell(colNum);
    }

    public IReadOnlyList<string> GetCards()
    {
        return _table.DataRange.Rows()
            .Select(r => DataCell(r, ColCard).GetString().Trim())
            .Where(c => !string.IsNullOrEmpty(c))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(c => c)
            .ToList();
    }

    public IReadOnlyList<Transaction> GetUnexported(string card)
    {
        var result = new List<Transaction>();

        foreach (var row in _table.DataRange.Rows())
        {
            var exportedCell = DataCell(row, ColExported);
            bool isExported = !exportedCell.IsEmpty() && exportedCell.DataType != XLDataType.Blank;
            if (isExported) continue;

            var cardVal = DataCell(row, ColCard).GetString().Trim();
            if (!string.Equals(cardVal, card, StringComparison.OrdinalIgnoreCase)) continue;

            result.Add(new Transaction
            {
                RowNumber   = row.WorksheetRow().RowNumber(),
                Date        = DataCell(row, ColDate).GetDateTime(),
                Payee       = DataCell(row, ColPayee).GetString().Trim(),
                Amount      = (decimal)DataCell(row, ColAmount).GetDouble(),
                Currency    = DataCell(row, ColCurrency).GetString().Trim().ToUpperInvariant(),
                Card        = cardVal,
                Category    = DataCell(row, ColCategory).GetString().Trim(),
                SubCategory = DataCell(row, ColSubCat).GetString().Trim(),
                Memo        = DataCell(row, ColMemo).GetString().Trim(),
            });
        }

        return result;
    }

    public void MarkExported(IEnumerable<Transaction> transactions, DateTime exportedAt)
    {
        if (!_colIndex.TryGetValue(ColExported.Trim(), out int colNum))
            throw new InvalidOperationException($"Column '{ColExported}' not found.");

        var ws = _workbook.Worksheet("Sheet1");
        foreach (var tx in transactions)
        {
            var cell = ws.Cell(tx.RowNumber, colNum);
            cell.Value = exportedAt;
            cell.Style.DateFormat.Format = "dd/MM/yyyy HH:mm";
        }

        _workbook.Save();
    }

    public void Dispose() => _workbook.Dispose();
}
