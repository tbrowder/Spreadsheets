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

Spreadsheet is ...

=head1 AUTHOR

Tom Browder <tom.browder@gmail.com>

=head1 COPYRIGHT AND LICENSE

Copyright 2020 Tom Browder

This library is free software; you can redistribute it and/or modify it under the Artistic License 2.0.

=end pod
