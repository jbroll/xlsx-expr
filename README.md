xlsx-expr
=========

Here is some code I've written to parse and evaluate formula in an Excel 2008 xml format spreadsheet. I generate some very complex reports in Excel format using perl (XLSX:Excel:Writer Using Perl to get Excel), so my xlsx output files do not have the values precomputed for each cell the way excel writes its output by default. To allow testing and verification of values that I'm pushing out to the client I need to be able to evaluate the formula that I've written. This is a rather long story but it starts with an example of using the Tcl Parser Tools from tcllib 1.13:

Down here in the guts of the thing are two Tcl gems.
vfs::zip
TAX: A Tiny API for XML
These two features make taking apart an MS Excel xlsx file pretty simple.
Here is the meat of the code. Its almost 500 lines so I'll just reference it here: http://rkroll.com/tclwiki/xlsx.tcl
Included are methods to open the xlsx file and parse the cell values, formula and formats from the xml into instance variable arrays in the xlsx object. The "=" method evaluates the value of the cell including following a formula right through all the references in the spreadsheet. Each formula value is cached and evaluated only once. Formula are evaluated by parsing them into AST format with parser tools and then executing the AST as a script. The result of the script is an expression suitable for expr, which is then called to obtain the cell's value. Just enough stuff is implemented here to support the syntax and functions in my spreadsheets, but extending this should be straight forward. I have evaluated workbooks with multiple worksheets, complex formula across thousands of cells with perfect agreement to the values that excel computes itself.
These Excel functions are currently supported:
MAX
MIN
COUNT
SUM
AND
DAYS360
VLOOKUP
IF - This is handled specially in the AST expansion to short circuit.
YEAR
ROUND
The parser is fed using tcl::chan::string which needs a small patch. The Allowance method is broken so I just added a "return" at the start of virtchannel_core/events.tcl:Allowance to disable any checking. This checking seemed like overkill to me anyway.
