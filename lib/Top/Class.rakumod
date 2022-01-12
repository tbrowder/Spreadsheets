unit module Top::Class;

class Foo {...};
class Bar is export {
    has Foo $.foo;
    submethod TWEAK {
        $!foo = Foo.new;
    }
}
class Foo is export {
    has $.id = 'Foo';
}
