#!/usr/bin/env raku

use Spreadsheet::Read:from<Perl5>;

my @f =
"../t/data/sample-security-sales.xlsx",
"../t/data/sample-security-sales.xls",
"../t/data/sample-security-sales.ods",
"../t/data/sample-security-sales.csv",
;

class WorkbookSet {...}
class Sheet {...}
class Workbook {
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

    method clone {
        # returns a copy of this Book object
    }
}
class WorkbookSet {
    #| an array of immutable input Workbook objects that can be written again under a new name
    has Workbook @.sources;
    has $.last-source-index = -1; # increment as source workbooks are added

    #| a hash of info on files read or written and their associated Workbook locations
    has %.files;

    #| an array of Workbook objects capable of being written
    has Workbook @.products;
    has $.last-product-index = -1; # increment as product workbooks are added

    method read(:$file!, :$debug) {
        # make sure the file isn't already in the hash
        my $fnam = $file.IO.basename;
        my $path = $file.IO.absolute;

        if %.files{$fnam}:exists {
            note "WARNING: File '$file' has already been read.";
            return;
        }
        if !$path.IO.f {
            note "FATAL: File '$file' cannot be read.";
            exit;
        }

        # figure out the correct workbook object to use
        %.files{$fnam}<path>         = $path;
        %.files{$fnam}<source-index> = ++$!last-source-index;
        my $wb = Workbook.new;
        @.sources.push: $wb;
        collect-file-data(:$path, :$wb, :$debug);
    }
}
class Cell {
    # should a Cell know its array position? just in case:
    has $.i; # row index, zero-based
    has $.j; # col index, zero-based

    has $.value;
    has $.read-format; # as reported by Spreadsheet::Read

    has $.format;

    method clone {
        # returns a copy of this Cell object
    }
}
class Row {
    has Cell @.cell; # an array of Cell objects

    method clone {
        # returns a copy of this Row object
    }
}
class Sheet {
    has Row @.row;      # an array of Row objects
    has %.colrow; # a hash indexed by Excel A1 label (col A, row 1)
    
    # check for and handle Excel colrow ids
    method add-colrow-hash($k, $v) {
        %.colrow; # a hash indexed by Excel A1 label (col A, row 1)
        if %.colrow{$k}:exists {
            note "WARNING: Excel A1 id '$k' is a duplicate";
        }
        else {
            %!colrow{$k} = $v;
        }
    }

    method dump-colrows {
        for %.colrow.keys.sort -> $k {
            my $v = %.colrow{$k};
            note "rolcow: $k, value: $v";
        }
    }

    method clone {
        # returns a copy of this Sheet object
    }
}

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
my $debug = 0;
for @*ARGS {
    when /^d/ { 
        $debug = 1; 
    }
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

my $c = WorkbookSet.new;
$c.read: :file($ifil);

exit;


#if $sheet > 1 and $n != 4 {
if $sheet > 1 and $ifil ~~ /:i csv/ {
    say "FATAL: Only one sheet in a csv file";
    exit;
}

my $book = ReadData $ifil;
my $ne = $book.elems;
say "\$book has $ne elements indexed from zero";
#exit;

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

        
        if $k ~~ /^ (<[A..Z]>+) (<[1..9]> <[0..9]>?) $/ {
            # collect the Excel A1 hashes
            my $col = ~$0;
            my $row = +$1;
            my $colrow = $col ~ $row.Str;

            note "DEBUG: found A1 Excel colrow id: '$k'" if $debug;
            if $t !~~ Str {
                note "WARNING: its value type is not Str it's: $t";
            }
            else {
                note "  DEBUG: with value: '$v'" if $debug;

                # need to confirm sheet num and its existence
                my $s = $wb.sheet[$sheet-1];

                # insert key and val in the sheet's %colrow hash
                $s.colrow{$k} = $v;
            }
        }
        elsif $k eq 'cell' {
            # collect the cell[col][row] values
        }

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

sub collect-file-data(:$path, Workbook :$wb!, :$debug) {
    my $pbook = ReadData $path; # arrray of hashes
    my $ne = $pbook.elems;
    say "\$book has $ne elements indexed from zero" if $debug;
    my %h = $pbook[0];
    collect-book-data %h, :$wb;

    # get all the sheet data
    for 1..^$ne -> $index {
        %h    = $pbook[$index];
        my $s = Sheet.new;
        $wb.sheet.push: $s;
        collect-sheet-data %h, :$index, :$s, :$debug;
    }
}

sub collect-book-data(%h, Workbook :$wb!, :$debug) {
    # Given the zeroth hash from Spreadsheet::Read and a 
    # Workbook object, collect the data for the workbook.
    constant %known-keys = set <
        error
        parser
        parsers
        quote
        sepchar
        sheet
        sheets
        type
        version
    >;

    for %h.kv -> $k, $v {
        note "WARNING: Unknown key '$k' in workbook meta data" unless %known-keys{$k}:exists;
    }
}

sub collect-sheet-data(%h, :$index, Sheet :$s!, :$debug) {
    # Given the sheet's original index, i, the ith hash 
    # from Spreadsheet::Read and a Sheet object, collect 
    # the data for the sheet.
    constant %known-keys = set <
        active
        attr
        cell
        indx
        label
        maxcol
        maxrow
        merged
        mincol
        minrow
        parser
    >;

    for %h.kv -> $k, $v {
        if $k ~~ /^ (<[A..Z]>+) (<[1..9]> <[0..9]>?) $/ {
            # check for and handle Excel colrow ids
            $s.add-colrow-hash: $k, $v;
            next;
        }

        note "WARNING: Unknown key '$k' in spreadsheet data" unless %known-keys{$k}:exists;
    }
}

sub collect-cell-data($cell, Sheet :$s!, :$debug) {
    # Given a cell array from Spreadsheet::Read and a 
    # Sheet object, collect the data for the sheet. In 
    # the process, convert the data into rows of cells 
    # with zero-based indexing.
}


