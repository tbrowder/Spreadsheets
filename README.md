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

AUTHOR
======

Tom Browder <tom.browder@gmail.com>

COPYRIGHT AND LICENSE
=====================

Copyright 2020 Tom Browder

This library is free software; you can redistribute it and/or modify it under the Artistic License 2.0.
