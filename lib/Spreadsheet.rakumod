unit class Spreadsheet:ver<0.0.1>:auth<cpan:TBROWDER>;

=begin pod

=head1 NAME

Spreadsheet - A universal spreadsheet reader/writer

=head1 SYNOPSIS

=begin code :lang<raku>

use Spreadsheet;
my $book = Spreadsheet.new;
$book.read: :file<myfile.cvs>, :has-header;
$book.write: :file<myfile.xlsx>;


=end code

=head1 DESCRIPTION

Spreadsheet is intended to be a reasonably universal spreadsheet
reader and writer for the formats shown below. It relies on some
well-tested Perl modules as well as a Raku module that wraps the
*libcsv* library available on Debian and other Linux distributions.

Its unique strength is a common set of classes to make spreadsheet
data use easy regardless of the file format being used.

=head2 Supported formats

=begin table
Read | Write | Notes
-----+-------+------
CSV  | *CSV  |
ODS  | *ODS  |
SXC  | *SXC  |
XLS  | *XLS  |
XLSX | XLSX  |
=end table

Note: Formats marked with an asterisk are not yet
implemented (NYI).

=head2 System requirements

=begin table
Perl modules                  | Debian package | Notes
---                           | ---            | ---
Spreadsheet::Read             | libspreadsheet-read-perl
Spreadsheet::ParseExcel       | libspreadsheet-parseexcel-perl
Spreadsheet::ParseXLSX        | libspreadsheet-parsexlsx-perl
Spreadsheet::ReadSXC          | libspreadsheet-readsxc-perl
Excel::Writer::XSLX           | libexcel-writer-xlsx-perl
Text::CSV                     | libtext-csv-perl
=end table

=head2 Design

This module is designed to treat data as a two-dimensional
array of data cells (row, column; zero indexed)
represented by a spreadsheet object. Multiple
spreadsheets can be children of a workbook object.

A spreadsheet may have the first row defined as a
header row with unique identifiers as keys to a
hash of each column.

Spreadsheet arrays may be acccessed in various ways to
suit the tastes of the user. For example, given a
spreadsheet $s:

=head3 Single cell (e.g., row 0, column 2)

=table
$s.cell(0,2)
$s.rowcol(0,2)
$s.colrow(2,0)
$s[0;2] | Raku syntax
$s<c1> | Excel syntax

=head3 Row of cells (a one-dimensional array)

=table
$s.row(0) | the entire row
$s[0;0..2] | row 0, columns 0 through 2
$s<1> | Excel syntax

=head3 Column of cells (a one-dimensional array)

=table
$s.col(0) | the entire column
$s[;0]
$s<a> | Excel syntax
$s[0..2;0] | column 0, rows 0 through 2
$s.col(0,0..2) | column 0, rows 0 through 2

=head3 Rectangular range of cells (a two-dimensional array)

=table
$s.rowcol(0..2,0..1) |
$s[0..2;0..1]
$s<a1:c2> | Excel syntax

=head2 Data model

The data model is based on the one described and used in Perl module
Spreadsheet::Read. Its data elements are used to populate the classes
described above (with adjustments to transform the 1-indexed rows and
columns to the zero-indexed rows and columns of this module).

=begin code
$book = [
    # Entry 0 is the overall control hash
    { sheets  => 2,
      sheet   => {
        "Sheet 1"  => 1,
        "Sheet 2"  => 2,
        },
      parsers => [ {
          type    => "xls",
          parser  => "Spreadsheet::ParseExcel",
          version => 0.59,
          }],
      error   => undef,
      },
    # Entry 1 is the first sheet
    { parser  => 0,
      label   => "Sheet 1",
      maxrow  => 2,
      maxcol  => 4,
      cell    => [ undef,
        [ undef, 1 ],
        [ undef, undef, undef, undef, undef, "Nugget" ],
        ],
      attr    => [],
      merged  => [],
      active  => 1,
      A1      => 1,
      B5      => "Nugget",
      },
    # Entry 2 is the second sheet
    { parser  => 0,
      label   => "Sheet 2",
      :
      :
=end code

=head1 AUTHOR

Tom Browder <tom.browder@gmail.com>

=head1 COPYRIGHT AND LICENSE

Copyright 2020 Tom Browder

This library is free software; you can redistribute it and/or modify it under the Artistic License 2.0.

=end pod

has %!data;

method read(:$file!, 
            :$debug, 
            :$index1,
           ) {
    use Spreadsheet::Read:from<Perl5>;

    # use Spreadsheet::Read to get all the data from the 
    # input file in its standard data format
    my $data = ReadData $file;
    say $data.gist;
}
