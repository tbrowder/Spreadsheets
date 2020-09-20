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
    # keys in the meta hash (book[0])
    #   with string values
    has $.quote   is rw = ''; # used for csv
    has $.sepchar is rw = ''; # used for csv
    has $.error   is rw = '';
    has $.sheets  is rw;      # number of sheets

    has $.parser  is rw;      # name of parser used
    has $.type    is rw;      # of the parser used: xlsx, xls, csv, etc.
    has $.version is rw;      # of the parser used
    #   with array or hash values
    has %.sheet   is rw;      # key: sheet name, value: index 1..N of N sheets

    # the following appears to be redundant and will be ignored on read iff it
    # only contains one element
    #has @.parsers is rw;      # array of parser pairs hashes, keys: name, type, version

    # convenience attrs
    has Sheet @.Sheet; # array of Sheet objects
    has $.basename = '';
    has $.path     = '';

    method dump(:$index!, :$debug) {
        say "DEBUG: dumping workbook index $index, file basename: {$.basename}";
        say "  == \%.sheet hash:";
        for %.sheet.keys.sort -> $k {
            my $v = %.sheet{$k};
            say "    '$k' => '$v'";
        }
    }

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

    method dump(:$debug) {
        my $ns = @.sources.elems;
        my $np = @.products.elems;
        my $s = $ns > 1 ?? 's' !! '';
        say "DEBUG: dumping WorkbookSet containing:";
        say "          $ns source workbook$s...";
        for @.sources.kv -> $i, $wb {
            $wb.dump: :index($i), :$debug;
        }
        $s = $np > 1 ?? 's' !! '';
        if $np {
            say "          and $np product workbook$s...";
        }
        else {
            say "          and no product workbooks.";
        }
    }

    method read(:$file!, :$debug) {
        # make sure the file isn't already in the hash
        my $basename = $file.IO.basename;
        my $path     = $file.IO.absolute;

        if %.files{$basename}:exists {
            note "WARNING: File '$file' has already been read.";
            return;
        }
        if !$path.IO.f {
            note "FATAL: File '$file' cannot be read.";
            exit;
        }

        # figure out the correct workbook object to use
        %.files{$basename}<path>         = $path;
        %.files{$basename}<source-index> = ++$!last-source-index;
        my $wb = Workbook.new: :$basename, :$path;
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
$c.read: :file($ifil), :debug;
if $debug {
    $c.dump;
}
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
    my $pbook = ReadData $path; # array of hashes
    my $ne = $pbook.elems;
    say "\$book has $ne elements indexed from zero" if $debug;
    my %h = $pbook[0];
    collect-book-data %h, :$wb, :$debug;

    # get all the sheet data
    for 1..^$ne -> $index {
        %h    = $pbook[$index];
        my $s = Sheet.new;
        $wb.Sheet.push: $s;
        collect-sheet-data %h, :$index, :$s, :$debug;
    }
}

sub collect-book-data(%h, Workbook :$wb!, :$debug) {
    # Given the zeroth hash from Spreadsheet::Read and a
    # Workbook object, collect the meta data for the workbook.

    constant %known-keys = [
        error    => 0,
        quote    => 0,
        sepchar  => 0,
        sheets   => 0,

        parser   => 0,
        type     => 0,
        version  => 0,

        parsers  => 0, # not used at the moment as it appears to be redundant
        sheet    => 0,
    ];

    say "DEBUG: collecting book meta data..." if $debug;
    for %h.kv -> $k, $v {
        say "  found key '$k'..." if $debug;
        note "WARNING: Unknown key '$k' in workbook meta data" unless %known-keys{$k}:exists;
        if $k eq 'error' {
            $wb.error = $v;
        }
        elsif $k eq 'parser' {
            $wb.parser = $v;
        }
        elsif $k eq 'quote' {
            $wb.quote = $v;
        }
        elsif $k eq 'sepchar' {
            $wb.sepchar = $v;
        }
        elsif $k eq 'sheets' {
            $wb.sheets = $v;
        }
        elsif $k eq 'type' {
            $wb.type = $v;
        }
        elsif $k eq 'version' {
            $wb.version = $v;
        }
        # special handling required
        elsif $k eq 'sheet' {
            $wb.sheet = get-wb-sheet-hash $v;
        }
        # special handling required
        elsif $k eq 'parsers' {
            # This appears to be redundant and will
            # be ignored as long as it only contains
            # one element. The one element is an anonymous
            # hash of three key/values (parser, type, version), all
            # which are already single-value attributes.
            my $ne = $v.elems;
            if $ne != 1 {
                die "FATAL: Expected one element but got $ne elements";
            }
        }
    }

    # ensure we have the parser, type, and version values as a sanity
    # check on our understanding of the read data format
    my $err = 0;
    if not $wb.parser {
        ++$err;
        note "WARNING: no 'parser' found in meta data";
    }
    if not $wb.type {
        ++$err;
        note "WARNING: no 'type' found in meta data";
    }
    if not $wb.version {
        ++$err;
        note "WARNING: no 'version' found in meta data";
    }
    if $err {
        note "POSSIBLE BAD READ OF FILE '$wb.path' PLEASE FILE AN ISSUE";
    }


}

sub get-wb-parsers-array($v) {
    my $t = $v.^name; # expect Perl5 Array
    my @a;
    my $val = $v // '';

    if $t ~~ /Array/ {
        if $val {
           for $val -> $v {
               my $t = $v.^name; # expect Perl5 Hash
               my $ne = $v.elems;
               note "DEBUG: element of parsers array is type: '$t'";
               note "       it has $ne element(s)";
               my $V = $v // '';
               @a.push: $V;
           }
        }
        else {
            note "array is empty or undefined";
        }
        return @a;
    }
    die "FATAL: Unexpected non-array type '$t'";
}

sub get-wb-sheet-hash($v) {
    my $t = $v.^name; # expect Perl5 Hash
    my %h;
    my $val = $v // '';

    if $t ~~ /Hash/ {
        if $val {
           for $val.kv -> $k, $v {
               %h{$k} = $v;
           }
        }
        return %h;
    }
    die "FATAL: Unexpected non-hash type '$t'";
}

sub collect-sheet-data(%h, :$index, Sheet :$s!, :$debug) {
    # Given the sheet's original index, i, the ith hash
    # from Spreadsheet::Read and a Sheet object, collect
    # the data for the sheet.
    constant %known-keys = [
        active   => 0,
        attr     => 0,
        cell     => 0,
        indx     => 0,
        label    => 0,
        maxcol   => 0,
        maxrow   => 0,
        merged   => 0,
        mincol   => 0,
        minrow   => 0,
        parser   => 0,
    ];

    for %h.kv -> $k, $v {
        if $k ~~ /^ (<[A..Z]>+) (<[1..9]> <[0..9]>?) $/ {
            # check for and handle Excel colrow ids
            $s.add-colrow-hash: $k, $v;
            next;
        }

        note "WARNING: Unknown key '$k' in spreadsheet data" unless %known-keys{$k}:exists;

        if $k eq 'active' {
            #$s. = $v;
        }
        elsif $k eq 'attr' {
            #$s. = $v;
        }
        elsif $k eq 'cell' {
            #$s. = $v;
        }
        elsif $k eq 'indx' {
            #$s. = $v;
        }
        elsif $k eq 'label' {
            #$s. = $v;
        }
        elsif $k eq 'maxcol' {
            #$s. = $v;
        }
        elsif $k eq 'maxrow' {
            #$s. = $v;
        }
        elsif $k eq 'merged' {
            #$s. = $v;
        }
        elsif $k eq 'mincol' {
            #$s. = $v;
        }
        elsif $k eq 'minrow' {
            #$s. = $v;
        }
        elsif $k eq 'parser' {
            #$s. = $v;
        }
    }
}

sub collect-cell-data($cell, Sheet :$s!, :$debug) {
    # Given a cell array from Spreadsheet::Read and a
    # Sheet object, collect the data for the sheet. In
    # the process, convert the data into rows of cells
    # with zero-based indexing.
}
