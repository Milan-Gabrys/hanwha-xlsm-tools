
# Hanwha XLSM Tools

## Krok 1b – odstranění ochran
In-memory úprava `.xlsm`: odstraní `<sheetProtection .../>` v určeném `sheetX.xml` a `<workbookProtection .../>` v `xl/workbook.xml`. Volba `-debug` rozbalí upravený ZIP do temp složky pro kontrolu.

### Použití
```bash
python src/hanwha_step1b_bytes.py "C:\cesta\vstup.xlsm" -debug
# nebo jiný sheet:
python src/hanwha_step1b_bytes.py "C:\cesta\vstup.xlsm" --sheet-xml "xl/worksheets/sheet1.xml" -debug
