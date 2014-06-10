
proc K { x y } { set x }
proc cat { file { newline {} } } {
    K [read {*}$newline [set fp [open $file]]] [close $fp]
}

namespace eval baseconvert {				# http://wiki.tcl.tk/16154
    variable chars "ABCDEFGHIJKLMNOPQRSTUVWXYZ"	;	# Modified!
	namespace export dec2base base2dec
}
proc baseconvert::dec2base {n b} {
    # algorithm found at http://www.rosettacode.org/wiki/Number_base_conversion#Python
    variable chars
	expr {$n == 0 ? 0
	    : "[string trimleft [dec2base [expr {$n/$b}] $b] 0][string index $chars [expr {$n%$b}]]"
	}
}
proc baseconvert::base2dec {n b} {
    variable chars
	set sum 0
	foreach char [split $n ""] {
	    set sum [expr {($sum * $b) + [string first $char $chars]}]
	}
    return $sum
}

proc colnum { let } { expr { [baseconvert::base2dec $let 26]+1 } }
proc collet { num } { baseconvert::dec2base [expr { $num-1 }] 26 }

oo::class create workbook {
    variable file worksheets sheet		\
	     SheetNumb SheetName SheetObj	\
	     sheets				\
    	     nstring Strings			\
	     stylestate nstyle Styles NumFormats

    constructor { File } {
    	set sheet    0
	set nstring  0
	set nstyle   0

	set file $File

	set NumFormats(0) {}

	vfs::zip::Mount $file /xlsx/$file
	set file /xlsx/$file

	if { [file exists $file/xl/sharedStrings.xml] } {
	    tax::parse [list [self] parser-strings] [cat $file/xl/sharedStrings.xml]
	}
	if { [file exists $file/xl/styles.xml] } {
	    tax::parse [list [self] parser-styles] [cat $file/xl/styles.xml]
	}
	tax::parse [list [self] parser] [cat $file/xl/workbook.xml]
	
	puts "Workbook Done"
    }

    method parser { tag cl self attr body } {
	switch $tag {
	 sheet {
	    incr sheet
	    set name [string map { &amp; & } [dict get $attr name]]
	    set SheetNumb($name) $sheet
	    set SheetName($sheet) $name

	    lappend sheets $name

	    set SheetObj($name) \
	        [worksheet create sheet$sheet [self] $sheet $file/xl/worksheets/sheet$sheet.xml]
	 }
	}
    }
    method parser-strings { tag cl self attr body } {
	switch $tag {
	 t { if { $cl } { return }
	     set Strings($nstring) $body
	     incr nstring
	 }
	}
    }
    method parser-styles { tag cl self attr body } {
	switch $tag {
	 numFmts { if { $cl } { return }
	     set stylestate numFmt
	 }
	 numFmt { if { $stylestate ne "numFmt" || $cl } { return }
	     set NumFormats([dict get $attr numFmtId]) [dict get $attr formatCode]
	 }
	 cellXfs { if { $cl } { return }
	     set stylestate cellXfs
	 }
	 xf { if { $stylestate ne "cellXfs" || $cl } { return }
	     set Styles($nstyle) $attr
	     incr nstyle
	 }
	}
    }
    method sheet { sheet args } { sheet$sheet {*}$args }
    method sheet2name { sheet } { return $SheetName($sheet) }
    method name2sheet {  name } { return $SheetNumb($name)  }
    method name2obj   {  name } { return $SheetObj($name)  }

    method string { n } { return $Strings($n) }
    method numfmt { n } {
	try { return [$NumFormats([dict get $Styles($n) numFmtId])] 
	} on error message {
	    return {}
	}
    }
    method sheets {} { return $sheets }
}

oo::class create worksheet {
	variable book name sheet state cell type style cells values formula numfmt

    constructor { Book Sheet File } {

	puts "$Book $Sheet $File"
	set  book  $Book
	set sheet $Sheet
	set state init

	set name [$book sheet2name $sheet]

	tax::parse [list [self] parser] [cat $File]

	puts Done
    }
    method parser { tag cl self attr body } {
	switch $tag {
	 sheetData { set state data }
	 
	 c { if { $state ne "data" || $cl } { return }
	    if { [dict exists $attr t] && [dict get $attr t] eq "s" } {
	     	set type s
	    } else {
	     	set type n
	    }
	    if { [dict exists $attr s] } {
	    	set style [dict get $attr s]
	    } else {
	    	set style 0
	    }

	    set cell  [dict get $attr r]
	    lappend cells $cell
	 }
	 v { if { $state ne "data" || $cl || $body eq "" } { return }
	    if { $type eq "s" } {
	        set values($cell) [$book string $body]
		set numfmt($cell) {}
	    } else {
	        if { $style } {
		    set numfmt($cell) [$book numfmt $style]
		} else {
		    set numfmt($cell) {}
		}
		set values($cell) $body
	    }
	 } 
	 f { if { $state ne "data" || $cl || $body eq "" } { return }
	     set  formula($cell) [string map { &amp; & &lt; < &gt; > } $body]
	 }
	}
    }
    method cells   {} { return $cells }
    method values  {} { return [array get values]  }
    method formula {} { return [array get formula] }
    method cell { cell } {
	set reply 0

	if { [info exists values($cell)] } {
	    if { [string is double $values($cell)] } {
	    	set reply [expr double($values($cell))]
	    } else {
		set reply $values($cell)
	    }
	} else {
	    if { [info exists  formula($cell)] } {
		try {
		    set reply [set values($cell) [my = $formula($cell)]]
		} on error message {
		    puts "while evaluating $name!$cell : $formula($cell) : $message"
		}
	    }
	}

	return $reply
    }
    method = { string } {
	#puts $string

    	set string [regsub -all {\m'?([^!(,="']+)'?!([A-Z]+)([0-9]+):([A-Z]+)([0-9]+)\M}  $string "{ { $book \\1 } \\2 \\3 \\4 \\5 }"]
    	set string [regsub -all {\m([A-Z]+)([0-9]+):([A-Z]+)([0-9]+)\M}  $string "{ [self] \\1 \\2 \\3 \\4 }"]
    	set string [regsub -all {\m([A-Z][A-Z]?[0-9]+)\M}  $string "cell(\"\\1\",\"[self]\")"]
    	set string [regsub -all {'?([^!(,="']+)'?!cell\("([^\"]+)",[^)]+\)} $string "cell(\"\\2\",sheet(\"$book\", \"\\1\"))"]
    	set string [regsub -all = $string ==]

	#puts $string

	return [expr $string]
    }
    method range-values { c1 r1 c2 r2 } {
        set reply {}

	set c1 [colnum $c1]
	set c2 [colnum $c2]
	for { set c $c1 } { $c <= $c2 } { incr c } {
	for { set r $r1 } { $r <= $r2 } { incr r } {
	    lappend reply [my cell [collet $c]$r]
	} }

	return $reply
    }
}

proc tcl::mathfunc::cell { cell sheet } {
    $sheet cell $cell
}
proc tcl::mathfunc::sheet { book sheet } {
    $book name2obj $sheet
}

proc range-values { args } {
    set reply {}
    foreach range $args {
        if { [llength $range] == 5 } {
	    set sheet [lindex $range 0]
	    if { [llength $sheet] == 2 } {
	        set sheet [[lindex $sheet 0] name2obj [lindex $sheet 1]]
	    }
	    lappend reply {*}[$sheet range-values {*}[lrange $range 1 4]]
	} else {
	    lappend reply $range
	}
    }

    set reply
}

proc tcl::mathfunc::Max { args } {
    set args [range-values {*}$args]
    set args [lassign $args max]

    foreach cell $args {
        if { [string is double $cell] } {
	    set max [expr max($max, $cell)]
	}
    }

    return $max
}
proc tcl::mathfunc::MIN { args } {
    set args [range-values {*}$args]
    set args [lassign $args min]

    foreach cell $args {
        if { [string is double $cell] } {
	    set min [expr min($min, $cell)]
	}
    }

    return $min
}

proc tcl::mathfunc::SUM { args } {
    set range [range-values {*}$args]
    set sum 0
    foreach cell $range {
        if { [string is double $cell] } {
	    set sum [expr {$sum+$cell}]
	}
    }
    set sum
}

proc tcl::mathfunc::DAYS360 { date1 date2 } {
    lassign [split [clock format [date $date1] -format "%m %d %Y"]] m1 d1 y1
    lassign [split [clock format [date $date2] -format "%m %d %Y"]] m2 d2 y2

    if { $d2 == 31 && $d1 >= 30 } {
	set d2 30
    }
    if { $d1 == 31 } {
	set d1 30
    }

    return [expr 360 * ($y2 - $y1) + 30 * ($m2 - $m1) + $d2 - $d1]
}

proc tcl::mathfunc::VLOOKUP { value range column } {
    lassign $range sheet c1 r1 c2 r2 
    set c3 [collet [expr [colnum $c1]+$column-1]]

    if { [string is double $value] } {
        set search -real
    } else {
    	set search -ascii
    }

    lindex                      [range-values [list $sheet $c3 $r1 $c3 $r2]]  \
	[lsearch -exact $search [range-values [list $sheet $c1 $r1 $c1 $r2]] $value]
}

proc xl2unx { date } {
    set date [expr int($date + 2440588 - 40587 + 15018)]
    clock scan $date -format %J
}
proc date { date } {
    if { [string is double $date] } {
        return [xl2unx $date]
    } else {
        return [clock scan $date -format "%m/%d/%Y"]
    }
}

proc tcl::mathfunc::YEAR { date } {
    clock format [date $date] -format %Y
}


proc tcl::mathfunc::YIELD { settle maturity coupon price value args } {
    set settle   [clock format [date $settle  ] -format "%Y%m%d"]
    set maturity [clock format [date $maturity] -format "%Y%m%d"]

    set yield [expr [exec yield $settle $maturity [expr $coupon*100] $price $value]/100.0]

    return $yield
}
proc tcl::mathfunc::ROUND { value n } {
    format %.${n}f $value
}
proc tcl::mathfunc::IF { bool true false } {
    if $bool { return $true
    } else   { return $false }
}

