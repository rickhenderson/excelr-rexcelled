# excelr-rexcelled
Part of a project to add R functionality into Excel. Eventually it will most likely become an add-in.

## List of Functions
* `read_csv(file, header = True, sep =",")`
* `plot()` - basic function to plot a simple chart on the data placed in the new worksheet created by read_csv

## Other Subs:
* `testReadCSV()` - a testing stub with multiple cases
* `Cleanup()` - sub to delete all worksheets in ThisWorkbook that don't have the name "Main"

### Log
**March 5, 2016**: Created the first iterations of read_csv() and got most functionality working.

## TODO
* Fix TAB separated file reading error
* Convert read_csv() to a function that returns an array
