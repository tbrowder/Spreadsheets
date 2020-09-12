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
CSV  | CSV   | uses `libcsv`
XLS  | XLS   |
XLSX | XLSX  |
ODS  | ODS   |
PSV  | PSV   |
=end table

=head1 AUTHOR

Tom Browder <tom.browder@gmail.com>

=head1 COPYRIGHT AND LICENSE

Copyright 2020 Tom Browder

This library is free software; you can redistribute it and/or modify it under the Artistic License 2.0.

=end pod
