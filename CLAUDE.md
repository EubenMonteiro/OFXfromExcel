# OFXfromExcel

## Project Purpose

A C# console app that reads an Excel-based financial transaction table and generates QIF files for import into Microsoft Money.

> **Note:** Microsoft Money exports QIF (not OFX) and imports both QIF and OFX. We target QIF because it is simpler and we have a validated sample export from Money to work against.

## Workflow

1. A button (to be aded to an app built in another session) triggers the C# routine
2. The routine reads all unprocessed rows from the transaction table
3. It generates an OFX file from those rows
4. It marks processed rows to prevent duplicate exports

## Excel Table Schema

Each row represents one financial transaction with these columns:

| Column       | OFX Field   | Notes                             |
|--------------|-------------|-----------------------------------|
| Date         | DTPOSTED    | Format: YYYYMMDD in OFX           |
| Payee        | NAME        | Counterpart / merchant name       |
| Amount       | TRNAMT      | Positive = credit, negative = debit |
| Category     | (memo part) | Used to build MEMO field          |
| Sub-Category | (memo part) | Used to build MEMO field          |
| Memo         | MEMO        | Free text; combined with category |

A "processed" flag column marks rows already exported to OFX.

## Multi-Currency Rules

The app supports two account types, selected at runtime:

### BRL accounts
- All transactions for the chosen card are exported, regardless of currency
- Non-BRL transactions are converted to BRL using a user-supplied exchange rate
- The original currency and amount are prepended to the Memo field: `EUR 5,10 - original memo`
- Output filename: `[Card]yyyyMMdd-HHmm.qif`

### Foreign currency accounts (EUR, USD, etc.)
- Only transactions matching the account currency are exported; others are skipped with a warning
- Amounts are exported as-is — no conversion, no memo prefix
- Output filename: `[Card] - [Currency] - yyyyMMdd-HHmm.qif`
- If the card name ends with a 3-letter currency code (e.g. `Avenue USD`), the currency is detected automatically and the user is not prompted

### Exchange rate caching
- Rates are grouped into 60-day windows anchored on transaction dates (looking backwards)
- The user is prompted once per currency per window; the rate is reused for all transactions in that window
- Rates are cached in `<ExcelFile>.rates.json` next to the workbook, storing `{currency, windowStart, windowEnd, rate}`

## OFX Format Reference (v1.x SGML)

Microsoft Money uses OFX 1.x SGML format (not XML). Key characteristics:

- Leaf elements have **no closing tags**: `<TRNAMT>-126.13` (no `</TRNAMT>`)
- Container/aggregate elements **require** matching closing tags: `<STMTTRN>...</STMTTRN>`
- Date format: `YYYYMMDD` (optionally `YYYYMMDDHHMMSS`)
- Amounts: decimal, no currency symbol, no thousands separator (e.g., `-1250.00`)
- FITID must be unique per transaction to prevent duplicate imports

### File Header (required, blank line after last header field)

```
OFXHEADER:100
DATA:OFXSGML
VERSION:102
SECURITY:NONE
ENCODING:USASCII
CHARSET:1252
COMPRESSION:NONE
OLDFILEUID:NONE
NEWFILEUID:NONE

```

### Minimal Bank Statement Structure

```
<OFX>
<SIGNONMSGSRSV1>
<SONRS>
<STATUS>
<CODE>0
<SEVERITY>INFO
</STATUS>
<LANGUAGE>POR
</SONRS>
</SIGNONMSGSRSV1>
<BANKMSGSRSV1>
<STMTTRNRS>
<TRNUID>1
<STMTRS>
<CURDEF>BRL
<BANKACCTFROM>
<BANKID>0000
<ACCTID>00000000
<ACCTTYPE>CHECKING
</BANKACCTFROM>
<BANKTRANLIST>
<DTSTART>20240101
<DTEND>20241231
<STMTTRN>
<TRNTYPE>DEBIT
<DTPOSTED>20240115
<TRNAMT>-250.00
<FITID>20240115001
<NAME>Supermercado Extra
<MEMO>Alimentacao:Supermercado - compras do mes
</STMTTRN>
<STMTTRN>
<TRNTYPE>CREDIT
<DTPOSTED>20240120
<TRNAMT>5000.00
<FITID>20240120001
<NAME>Empresa XYZ
<MEMO>Salario:Janeiro
</STMTTRN>
</BANKTRANLIST>
<LEDGERBAL>
<BALAMT>4750.00
<DTASOF>20241231
</LEDGERBAL>
</STMTRS>
</STMTTRNRS>
</BANKMSGSRSV1>
</OFX>
```

### TRNTYPE Values

`CREDIT`, `DEBIT`, `DEP`, `INT`, `DIV`, `FEE`, `SRVCHG`, `ATM`, `POS`, `XFER`, `CHECK`, `PAYMENT`, `CASH`, `DIRECTDEP`, `DIRECTDEBIT`, `REPEATPMT`, `OTHER`

### MEMO Field Convention

Build MEMO as: `Category:Sub-Category - free memo text`

For foreign-currency transactions: `EUR 5,10 Category:Sub-Category - free memo text`

## Excel Table Schema

File: `Recibos para o Money.xlsx`, Sheet: `Sheet1`, Table name: `Money` (A:I)

| Column | Header | QIF field | Notes |
|--------|--------|-----------|-------|
| A | Data | D | Date posted |
| B | Estabelecimento | P | Payee |
| C | Valor | T | Amount — positive in Excel = debit; code multiplies by -1 |
| D | Moeda | — | Currency code (BRL default; EUR/USD etc. triggers conversion) |
| E | Cartão | — | Card name; user selects which card to export each run |
| F | Categoria | L | QIF category |
| G | Sub-Cat | L | QIF subcategory (appended as `Category:Subcategory`) |
| H | Memo | M | Free text memo |
| I | Exportado em | — | Datetime stamp written by the app after export; null = unprocessed |

## QIF Format Reference

```
!Type:Bank
D26/06'2026
T-999.99
PPayee name
MMemo text
LCategory:Subcategory
^
```

- Date format: `DD/MM'YYYY`
- Amount: negative = debit, positive = credit
- `^` separates records
- Multi-currency memo prefix: `EUR 5,10 - original memo text`

## Project Structure

```
OFXfromExcel/
├── CLAUDE.md
├── Recibos para o Money - cópia teste.xlsx   # sample workbook (git-tracked)
├── Teste Geração OFX com Claude 1.qif        # sample QIF export from Money
└── QifFromExcel/                             # C# console app (.NET 8)
    ├── QifFromExcel.csproj                   # ClosedXML dependency
    ├── Program.cs                            # Entry point; CONFIGURATION constants at top
    ├── ExcelReader.cs                        # Reads Money table, marks exported rows
    ├── QifGenerator.cs                       # Builds QIF string from transactions
    ├── ExchangeRateCache.cs                  # Prompts for rates, caches 60 days in .rates.json
    └── Models/
        └── Transaction.cs                    # Domain model for one row
```

## Configuration (Program.cs top)

Two constants control file paths — change these to match the production file locations:

```csharp
const string ExcelFilePath = @"C:\Users\euben\...\Recibos para o Money.xlsx";
const string OutputFolder  = @"C:\Users\euben\...\Gerando OFX";
```

## Development Notes

- Runtime: .NET 8 console app — run with `dotnet run` or publish as a single `.exe`
- Excel library: ClosedXML (MIT licence) — no licence prompt on commercial use
- Exchange rates cached in `<ExcelFile>.rates.json` next to the workbook; refreshed after 60 days
- Output filename pattern: `[Card]yyyyMMdd-HHmm.qif`
- The same timestamp written to the output filename is stamped into column I of the exported rows
