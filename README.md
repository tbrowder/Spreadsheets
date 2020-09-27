[![Actions Status](https://github.com/tbrowder/Spreadsheets/workflows/test/badge.svg)](https://github.com/tbrowder/Spreadsheets/actions)

NAME
====

Spreadsheet - A universal, multiple spreadsheet reader/writer

SYNOPSIS
========

```raku
use Spreadsheets;
my $book = Spreadsheets.new;
$book.read: :file<myfile.cvs>, :has-header;
$book.write: :file<myfile.xlsx>;
```

DESCRIPTION
===========

Spreadsheet is intended to be a reasonably universal spreadsheet reader and writer for the formats shown below. It relies on some well-tested Perl modules as well as a Raku module that wraps the *libcsv* library available on Debian and other Linux distributions.

Its unique strength is a common set of classes to make spreadsheet data use easy regardless of the file format being used.

Supported formats
-----------------

<table class="pod-table">
<thead><tr>
<th>Read</th> <th>Write</th> <th>Notes</th>
</tr></thead>
<tbody>
<tr> <td>CSV</td> <td>*CSV</td> <td></td> </tr> <tr> <td>ODS</td> <td>*ODS</td> <td></td> </tr> <tr> <td>SXC</td> <td>*SXC</td> <td></td> </tr> <tr> <td>XLS</td> <td>*XLS</td> <td></td> </tr> <tr> <td>XLSX</td> <td>XLSX</td> <td></td> </tr>
</tbody>
</table>

Note: Formats marked with an asterisk are not yet implemented (NYI). The author does not intend to expend any effort on writing those formats unless he gets a Pull Request (PR) which provides such a capability.

System requirements
-------------------

<table class="pod-table">
<thead><tr>
<th>Perl modules</th> <th>Debian package</th> <th>Notes</th>
</tr></thead>
<tbody>
<tr> <td>Spreadsheet::Read</td> <td>libspreadsheet-read-perl</td> <td></td> </tr> <tr> <td>Spreadsheet::ParseExcel</td> <td>libspreadsheet-parseexcel-perl</td> <td></td> </tr> <tr> <td>Spreadsheet::ParseXLSX</td> <td>libspreadsheet-parsexlsx-perl</td> <td></td> </tr> <tr> <td>Spreadsheet::ReadSXC</td> <td>libspreadsheet-readsxc-perl</td> <td></td> </tr> <tr> <td>Text::CSV</td> <td>libtext-csv-perl</td> <td></td> </tr> <tr> <td>Excel::Writer::XSLX</td> <td>libexcel-writer-xlsx-perl</td> <td></td> </tr>
</tbody>
</table>

Design
------

This module is designed to treat data as a two-dimensional array of data cells (row, column; zero indexed), commonly referred to as a 'spreadsheet', represented by a Sheet object. Multiple spreadsheets can be children of a Workbook object which is modeled after an Excel XLSX file (known as a workbook). Finally, a WorkbookSet object can have multiple Workbook objects as children.

A spreadsheet may have the first row defined as a header row with unique identifiers as keys to a hash of each column.

Spreadsheet arrays may be acccessed in various ways to suit the tastes of the user. For example, given a spreadsheet $s:

### Single cell (e.g., row 0, column 2)

<table class="pod-table">
<tbody>
<tr> <td>$s.cell(0,2)</td> <td></td> </tr> <tr> <td>$s.rowcol(0,2)</td> <td></td> </tr> <tr> <td>$s.colrow(2,0)</td> <td></td> </tr> <tr> <td>$s[0;2]</td> <td>Raku syntax</td> </tr> <tr> <td>$s&lt;c1&gt;</td> <td>Excel syntax</td> </tr>
</tbody>
</table>

### Row of cells (a one-dimensional array)

<table class="pod-table">
<tbody>
<tr> <td>$s.row(0)</td> <td>the entire row</td> </tr> <tr> <td>$s[0;0..2]</td> <td>row 0, columns 0 through 2</td> </tr> <tr> <td>$s&lt;1&gt;</td> <td>Excel syntax</td> </tr>
</tbody>
</table>

### Column of cells (a one-dimensional array)

<table class="pod-table">
<tbody>
<tr> <td>$s.col(0)</td> <td>the entire column</td> </tr> <tr> <td>$s[;0]</td> <td></td> </tr> <tr> <td>$s&lt;a&gt;</td> <td>Excel syntax</td> </tr> <tr> <td>$s[0..2;0]</td> <td>column 0, rows 0 through 2</td> </tr> <tr> <td>$s.col(0,0..2)</td> <td>column 0, rows 0 through 2</td> </tr>
</tbody>
</table>

### Rectangular range of cells (a two-dimensional array)

<table class="pod-table">
<tbody>
<tr> <td>$s.rowcol(0..2,0..1)</td> <td></td> </tr> <tr> <td>$s[0..2;0..1]</td> <td></td> </tr> <tr> <td>$s&lt;a1:c2&gt;</td> <td>Excel syntax</td> </tr>
</tbody>
</table>

Data model
----------

The data model is based on the one described and used in Perl module Spreadsheet::Read. Its data elements are used to populate the classes described above (with adjustments to transform the 1-indexed rows and columns to the zero-indexed rows and columns of this module).

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

AUTHOR
======

Tom Browder <tom.browder@gmail.com>

COPYRIGHT AND LICENSE
=====================

Copyright 2020 Tom Browder

This library is free software; you can redistribute it and/or modify it under the Artistic License 2.0.

