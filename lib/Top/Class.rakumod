unit module Top::Class;

class Foo {...}
class Baz {...}
class Bar is export {
    has Foo $.foo;
    submethod TWEAK {
        $!foo = Foo.new;
    }
}
class Foo is export {
    has $.id = 'Foo';
}
class Baz is export {
    has $.id = 'Baz';
}
