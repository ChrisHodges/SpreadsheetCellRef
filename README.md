# SpreadsheetCellRef
A [.Net struct](https://docs.microsoft.com/en-us/dotnet/csharp/language-reference/builtin-types/struct) representing the address of a cell in a spreadsheet.

## Usage
Create a `CellRef` from a string:

``` csharp
var cellRefA1 = new CellRef("A1");

```

or from 1-based row and column integers:

``` csharp
var cellRefA1FromInts = new CellRef(1,1);
```

The 2 `cellRef`s created above will be equal:

``` csharp
var cellRefA1 = new CellRef("A1");
var cellRefA1FromInts = new CellRef(1,1);
if (cellRefA1 == cellRefA1FromInts) {
    Console.WriteLine('Cell refs are equal');
}
```

A `cellRef`s row and column values can be accessed from the `Row`, `Column` and `ColumnNumber` properties:

``` csharp
var cellRef = new CellRef("AB32");
Console.WriteLine(cellRef.Row); //Outputs '32'
Console.WriteLine(cellRef.Column); //Outputs 'AB'
Console.WriteLine(cellRef.ColumnNumber); //Outputs '28'
```

Calling `ToString()` on a cellRef outputs the cell address in ColumnString-RowInt format:

``` csharp
var cellRef = new CellRef("E5");
Console.WriteLine(cellRef.ToString()); //Outputs 'E5';
```

### Adding Rows and Columns:
You can use the `AddRows(int n)` and `AddColumns(int n)` methods to return a new `CellRef` **n** rows or columns away from the original:

``` csharp
var cellRefA1 = new CelRef("A1");
var cellRefA6 = cellRefA1.AddRows(5);
var cellRefC1 = cellRefA1.AddColumns(2);
```

### Ranges
Create a range of `cellRef`s using the static `Range` method. Cells are returned in left-to-right then top-to-bottom order:

``` csharp
IEnumerable<CellRef> range = CelRef.Range("A1","B2");
Console.WriteLine(range[0]); //Outputs 'A1'
Console.WriteLine(range[1]); //Outputs 'B1'
Console.WriteLine(range[2]); //Outputs 'A2'
Console.WriteLine(range[3]); //Outputs 'B2'
```

### Converting Column Strings to Column Numbers (and vice versa)
You can convert a string representation of a column name to an integer using the static method `CellRef.ColumnNameToNumber`. To convert an integer to a string representation you can use `CellRef.ColumnNameToNumber`:

``` csharp
Console.WriteLine(CellRef.ColumnNameToNumber("AA")); //Outputs '27'
Console.WriteLine(CellRef.NumberToColumnName(27)); //Outputs 'AA'
```
