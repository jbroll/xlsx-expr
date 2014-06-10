namespace eval tax {}

# ::tax::__cleanprops -- Clean parsed XML properties
#
#	This command cleans parsed XML properties by removing the
#	trailing slash and replacing equals by spaces so as to produce
#	a list that is suitable for an array set command.
#
# Arguments:
#	props	Parsed XML properties
#
# Results:
#	Return an event list that is suitable for an array set
#
# Side Effects:
#	None.
proc ::tax::__cleanprops { props } {
    set name {([A-Za-z_:]|[^\x00-\x7F])([A-Za-z0-9_:.-]|[^\x00-\x7F])*}
    set attval {"[^"]*"|'[^']*'|\w}; # "... Makes emacs happy
    return [regsub -all -- "($name)\\s*=\\s*($attval)" \
	[regsub "/$" $props ""] "\\1 \\4"]
}

# ::tax::parse -- Low-level 10 lines magic parser
#
#	This procedure is the core of the tiny XML parser and does its
#	job in 10 lines of "hairy" code.  The command will call the
#	command passed as an argument for each XML tag that is found
#	in the XML code passed as an argument.  Error checking is less
#	than minimum!  The command will be called with the following
#	respective arguments: name of the tag, boolean telling whether
#	it is a closing tag or not, boolean telling whether it is a
#	self-closing tag or not, list of property (array set-style)
#	and body of tag, if available.
#
# Arguments:
#	cmd	Command to call for each tag found.
#	xml	String containing the XML to be parsed.
#	start	Name of the pseudo tag marking the beginning/ending of document
#
# Results:
#	None.
#
# Side Effects:
#	None.
proc ::tax::parse {cmd xml {start docstart}} {
    regsub -all \{ $xml {\&ob;} xml
	regsub -all \} $xml {\&cb;} xml
    set exp {<(/?)([^\s/>]+)\s*([^>]*)>}
    set sub "\}\n$cmd {\\2} \[expr \{{\\1} ne \"\"\}\] \[regexp \{/$\} {\\3}\] \
    \[::tax::__cleanprops \{\\3\}\] \{"
    regsub -all $exp $xml $sub xml
    eval "$cmd {$start} 0 0 {} \{$xml\}"
    eval "$cmd {$start} 1 0 {} {}"
}
