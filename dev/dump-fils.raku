#!/usr/bin/env raku

use Text::Utils :normalize-string;

use lib <../lib>;

use Spreadsheets::Utils;

my @f =
"../t/data/sample-security-sales.xlsx",
"../t/data/sample-security-sales.xls",
"../t/data/sample-security-sales.ods",
"../t/data/sample-security-sales.csv",
"../t/data/mytest.csv",
"../t/data/senior-center-schedule.xlsx",
"../t/data/tmp-sto/senior-center-schedule-orig-buggy.xlsx",
;

my $sheet = 0;
if !@*ARGS.elems {
    say qq:to/HERE/;
    Usage: {$*PROGRAM.basename} 1|2|3|4|5|6|7  [s1 s2] [debug]

    Uses the Perl module Spreadsheet::Read and
    dumps the data from the selected file number:
    HERE
    my $n = 0;
    for @f -> $f {
        ++$n;
        say "  $n. {$f.IO.basename}";
    }
    say();
    exit;
}

my $n;
my $debug = 0;
for @*ARGS {
    when /^d/ {
        $debug = 1;
    }
    when /s(1|2)/ {
        $sheet = +$0;
    }
    when /(1|2|3|4|5|6|7)/ {
        $n = +$0 - 1
    }
    default {
        say "FATAL: Unhandled arg '$_'";
        exit;
    }
}

my $ifil = @f[$n];

#use Spreadsheet::Read:from<Perl5>;

#my $c = Spreadsheets::Classes::WorkbookSet.new;
my $c = WorkbookSet.new;

#=finish

$c.read: :file($ifil), :$debug;
if $debug {
    $c.dump;
}
say "DEBUG early exit after dump"; exit;

if $sheet > 1 and $ifil ~~ /:i csv/ {
    say "FATAL: Only one sheet in a csv file";
    exit;
}


=finish

# note the following read line is critical for interpreting
# the input data
my $book = Spreadsheet::Read::ReadData $ifil,
    :attr(1),
    #:clip(1),
    #:strip(3)
    ;

my $ne = $book.elems;
say "\$book has $ne elements indexed from zero";
#exit;

my @rows = $book[1].rows;
say "DEBUG: \@rows.gist:";
say @rows.gist;
exit;

my %h = $book[$sheet];
say "Dumping hash in \$book[$sheet]:";
my $wb = Workbook.new;
for 1..$ne {
    my $s = Sheet.new;
    $wb.sheet.push: $s;
}

=begin comment
my $s1 = Sheet.new;
my $s2 = Sheet.new;
$wb.sheet.push: $s1;
$wb.sheet.push: $s2;
=end comment

#dump-hash %h, :$debug;
say "\$book has $ne elements indexed from zero";

my $idx = 0;
for $book[1..*] -> $arr {
    ++$idx;
    my $n = $arr.elems;
    say "\$book[$idx] has $n elements";
}
for $wb.sheet -> $s {
    $s.dump-colrows;
}

#$s1.dump-colrows;
#$s2.dump-colrows;


exit;


%h = $book[1];
say "Dumping hash in \$book[1]:";
dump-hash %h;
