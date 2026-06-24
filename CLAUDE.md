# OFXfromExcel

## Project Purpose

A C# routine integrated into an Excel-based financial transaction app that generates OFX files from an Excel table for import into Microsoft Money (or compatible apps).

## Workflow

1. A button in Excel triggers the C# routine
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

- Default currency: **BRL**
- If a row has a non-BRL currency (e.g., EUR, USD):
  - Prepend `"EUR 5,10 "` (or relevant currency + original amount) to the MEMO field
  - Convert the amount to BRL at a user-supplied exchange rate
  - Exchange rates are prompted once per ~60-day window and cached; the user is not re-prompted within the same window

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

## Project Structure

```
OFXfromExcel/
├── CLAUDE.md               # This file
├── OFXfromExcel.sln        # Visual Studio solution (to be created)
├── OFXfromExcel/           # C# class library project
│   ├── OFXGenerator.cs     # Core OFX file generation logic
│   ├── ExcelReader.cs      # Reads and marks rows in Excel table
│   ├── CurrencyConverter.cs # Exchange rate prompting and caching
│   └── Models/
│       ├── Transaction.cs  # Domain model for one row
│       └── ExchangeRateWindow.cs
└── Tests/                  # Unit test project (optional)
```

## Development Notes

- Target: .NET Framework or .NET 8 (confirm with Excel interop requirements)
- Excel integration: likely via Excel-DNA or VSTO; the C# code is invoked from VBA or a ribbon button
- OFX encoding: write files as Windows-1252 (CHARSET:1252) to match the header
- FITID generation: use `{date:yyyyMMdd}{row_index:000}` to guarantee uniqueness within a file
- The generated `.ofx` file should be saved to a user-chosen path, then the user manually imports it into Microsoft Money

## Pending Verification

- [ ] Confirm exact OFX structure Microsoft Money expects by examining a sample export
- [ ] Confirm TRNTYPE mapping from Category/Sub-Category values in the Excel table
- [ ] Confirm how the "processed" flag is stored in Excel (column value, cell color, etc.)
- [ ] Confirm Excel interop approach (Excel-DNA add-in vs. standalone .exe vs. VSTO)
