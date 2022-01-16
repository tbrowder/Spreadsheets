#!/usr/bin/env raku


my $na = 3;
my $nb = 4;

SHEET: for 0..^$na -> $a {
    if $a == 1 {
        die "WARNING: bad sheet $a, skipping it entirely"; # fail, recover at $a = 2
        CATCH { default { say .Str; next SHEET; } }
    }
    say "== sheet $a";

    CELL: for 0..^$nb -> $b {
        if $a == 2 and $b == 2 {
            die "WARNING: sheet $a, bad cell $b"; # fail, recover at $b = 4
            CATCH { default { say .Str; next CELL; } }
        }
        say "  == cell $b";
    }
}

