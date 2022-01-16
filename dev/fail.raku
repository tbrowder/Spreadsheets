#!/usr/bin/env raku

# Simulating a spreadsheet sheet/cell reader:
my $ns = 3; # number of sheets
my $nc = 4; # number of cells per sheet
SHEET: for 1..$ns -> $s {
    if $s == 2 {
        # a sheet failure, recover at $s = 3
        die "WARNING: bad sheet $s, skipping it entirely";
        CATCH { default { say .Str; next SHEET; } }
    }
    say "== sheet $s";

    CELL: for 1..$nc -> $c {
        if $s == 3 and $c == 3 {
            # a cell failure, recover at $c = 4
            die "WARNING: sheet $s, bad cell $c";
            CATCH { default { say .Str; next CELL; } }
        }
        say "  == cell $c";
    }
}
