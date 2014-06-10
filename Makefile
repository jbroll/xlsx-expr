
PT=/home/john/src/tcllib-1.13/modules/pt/pt

parser : xls-parser.peg
	$(PT) generate snit -class xls-parser  -name xls-parser xls-parser.tcl peg xls-parser.peg
	sed -e s/PACKAGE/xls-parser/ < xls-parser.tcl > tmp
	mv tmp xls-parser.tcl 


xxx:
	#$(PT) generate oo -class xls-expr  -name xls-expr xls-expr.tcl peg xls-expr.peg
	#sed -e s/OO/TclOO/ < xls-expr.tcl > tmp
	#mv tmp xls-expr.tcl
