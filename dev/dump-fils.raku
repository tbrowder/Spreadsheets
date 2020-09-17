#!/usr/bin/env raku

my @f =
"../t/data/sample-security-sales.xlsx",
"../t/data/sample-security-sales.xls",
"../t/data/sample-security-sales.ods",
"../t/data/sample-security-sales.csv",
;

class Sheet {...}
class Book {
    has $.quote   = ''; # used for csv
    has $.sepchar = ''; # used for csv
    has $.error   = '';
    has $.parser;      # name of parser used
    has @.parsers;     # array of parser hashes, keys: name, type, version
    has $.num-sheets;
    has $.type;        # of the parser used: xlsx, xls, csv, etc.
    has $.version;     # of the parser used
    has %.sheets;      # key: sheet name, value: index 1..N of N sheets
    has Sheet @.sheet; # array of Sheet objects
}
class Cell {
    has $.value;
    has $.format;
}
class Row {
    has Cell @.cell; # an array of Cell objects
}
class Sheet {
    has Row @.row; # an array of Row objects
}

use Spreadsheet::Read:from<Perl5>;

my $sheet = 0;
if !@*ARGS.elems {
    say qq:to/HERE/;
    Usage: {$*PROGRAM.basename} 1|2|3|4  [s1 s2]
    
    Uses the Perl  module Spreadsheet::Read and 
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
for @*ARGS {
    when /s(1|2)/ { 
        $sheet = +$0; 
    }
    when /(1|2|3|4)/ { 
        $n = +$0 - 1 
    }
    default {
        say "FATAL: Unhandled arg '$_'";
        exit;
    }
}

my $ifil = @f[$n];
#if $sheet > 1 and $n != 4 {
if $sheet > 1 and $ifil ~~ /:i csv/ {
    say "FATAL: Only one sheet in a csv file";
    exit;
}

my $book = ReadData $ifil;
my $ne = $book.elems;
say "\$book has $ne elements indexed from zero";

my %h = $book[$sheet];
say "Dumping hash in \$book[$sheet]:";
dump-hash %h;
say "\$book has $ne elements indexed from zero";

my $idx = 0;
for $book[1..*] -> $arr {
    ++$idx;
    my $n = $arr.elems;
    say "\$book[$idx] has $n elements";
}

exit;


%h = $book[1];
say "Dumping hash in \$book[1]:";
dump-hash %h;

#### subroutines ####
sub dump-array(@a, :$level is copy = 0, :$debug) {
    my $sp = $level ?? '  ' x $level !! '';
    for @a.kv -> $i, $v {
        my $t = $v.^name;
        say "$sp index $i, value type: $t";
        if $t ~~ /Hash/ {
            dump-hash $v, :level(++$level), :$debug;
        }
        elsif $t ~~ /Array/ {
            # we may have an undef array
            my $val = $v // '';
            if $val {
                dump-array $v, :level(++$level), :$debug;
            }
            else {
                say "$sp   (undef array)";
            }
        }
        else {
            my $s = $v // '';
            say "$sp   value: '$s'";
        }
    }
}

sub dump-hash(%h, :$level is copy = 0, :$debug) {
    my $sp = $level ?? '  ' x $level !! '';
    for %h.keys.sort -> $k {
        my $v = %h{$k} // '';
        my $t = $v.^name;
        say "$sp key: $k, value type: $t";
        if $t ~~ /Hash/ {
            dump-hash $v, :level(++$level), :$debug;
        }
        elsif $t ~~ /Array/ {
            # we may have an undef array
            my $val = $v // '';
            if $val {
                dump-array $v, :level(++$level), :$debug;
            }
            else {
                say "$sp   (undef array)";
            }
        }
        else {
            my $s = $v // '';
            say "$sp   value: '$s'";
        }
    }
}

