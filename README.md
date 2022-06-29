# Convert data from ODS to USV

Convert a data spreadsheet from the format OpenDocument Spreadsheet (ODS) into the format Unicode Separated Values (USV).

Example:

```sh
$ target/debug/convert-data-from-ods-into-usv test/example.ods
worksheet name: Sheet1, range: Range { start: (0, 0), end: (1, 2), inner: [String("A1"), String("B1"), String("C1"), String("A2"), String("B2"), String("C2")] }, height: 2, width: 3
row len: 3
data: String("A1")
data: String("B1")
data: String("C1")
row len: 3
data: String("A2")
data: String("B2")
data: String("C2")
```