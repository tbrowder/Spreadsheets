#!/usr/bin/env raku

use lib <../lib>;
use Spreadsheet;

my $book = Spreadsheet.new;

my @f =
"../t/data/sample-security-sales.xlsx",
"../t/data/sample-security-sales.xls",
"../t/data/sample-security-sales.ods",
"../t/data/sample-security-sales.csv",
;

if !@*ARGS.elems {
    say qq:to/HERE/;
    Usage: {$*PROGRAM.basename} 1|2|3|4 
    
    Dumps the data from the selected file number:
    HERE
    my $n = 0;
    for @f -> $f {
        ++$n;
        say "  $n. {$f.IO.basename}";
    }
    exit;
}

my $n;
for @*ARGS {
    when /(1|2|3|4)/ { 
        $n = +$0 - 1 
    }
    default {
        say "FATAL: Unhandled arg '$_'";
        exit;
    }
}

$book.read: @f[$n];
say $book.gist;
say "The above data were in file '@f[$n]'";

