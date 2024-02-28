# Excel VBA CRM for a retail grocery store

Spreadsheets:
1) Products (where barcodes and product identity are inserted).
2) Received Products (IN) table (for every new product that comes in. Requires to specify "Quantity" and "Unit price bought").
3) Sales (OUT) table (where every sale is registered. It also requires specifying the quantity for any product sold).
4) Inventory table (every barcode product is registered once. Columns automatically update following "Received Products" & "Sales").
5) Income Statement Sheet (automatically updates). You can select the desired date.

Excel formulas used:
1) VLOOKUP().
2) SUMIF().
3) SUMIFS().
4) IF().
5) ISBLANK().
6) SUM().
