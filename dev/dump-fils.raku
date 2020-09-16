#!/usr/bin/env raku

my @f =
"../t/data/sample-security-sales.xlsx",
"../t/data/sample-security-sales.xls",
"../t/data/sample-security-sales.ods",
"../t/data/sample-security-sales.csv",
;

use Spreadsheet::Read:from<Perl5>;

if !@*ARGS.elems {
    say qq:to/HERE/;
    Usage: {$*PROGRAM.basename} 1|2|3|4 
    
    Uses the Perl  module Spreadsheet::Read and 
    dumps the data from the selected file number:
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

my $book = ReadData @f[$n];

my %h = $book[0];
say "Dumping hash in \$book[0]:";
dump-hash %h;
#exit;

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
            dump-array $v, :level(++$level), :$debug;
        }
        else {
            my $s = $v // '';
            say "$sp   value: '$s'";
        }
    }
}

