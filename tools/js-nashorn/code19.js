var Callable = Java.type("java.util.concurrent.Callable");
 
var FooCallable = Java.extend(Callable, {
  call: function() {
    return "Foo";
  }
});

var foobar = new FooCallable({
  call: function() {
    return "FooBar";
  }
});
 
// 'FooBar'
print(foobar.call());
 
// 'true'
print(foo.getClass() === foobar.getClass());
