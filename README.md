NAME
====

Spreadsheet - A universal spreadsheet reader/writer

SYNOPSIS
========

```raku
use Spreadsheet;
my $book = Spreadsheet.new;
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
<tr> <td>CSV</td> <td>CSV</td> <td>uses `libcsv`</td> </tr> <tr> <td>XLS</td> <td>XLS</td> <td></td> </tr> <tr> <td>XLSX</td> <td>XLSX</td> <td></td> </tr> <tr> <td>ODS</td> <td>ODS</td> <td></td> </tr> <tr> <td>PSV</td> <td>PSV</td> <td></td> </tr>
</tbody>
</table>

AUTHOR
======

Tom Browder <tom.browder@gmail.com>

COPYRIGHT AND LICENSE
=====================

Copyright 2020 Tom Browder

This library is free software; you can redistribute it and/or modify it under the Artistic License 2.0.

