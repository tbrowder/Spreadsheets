unit class Top;


=finish

use Top::Class;

has Bar $.bar;
has Foo $.foo;

submethod TWEAK {
    $!bar = Bar.new;
    $!foo = Foo.new;
}

