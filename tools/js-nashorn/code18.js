var Callable = Java.type("java.util.concurrent.Callable");
 
var FooCallable = Java.extend(Callable, {
  call: function() {
    return "Foo";
  }
});
 
var BarCallable = Java.extend(Callable, {
  call: function() {
    return "Bar";
  }
});
 
var foo = new FooCallable();
var bar = new BarCallable();
 
// 'false'
print(foo.getClass() === bar.getClass());
 
print(foo.call());
print(bar.call());
