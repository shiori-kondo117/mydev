var iterator = new java.util.Iterator({
  i: 0,
  hasNext: function() {
    return this.i < 10;
  },
  next: function() {
    return this.i++;
  }
});
 
print(iterator instanceof Java.type("java.util.Iterator"));
while (iterator.hasNext()) {
  print("-> " + iterator.next());
}
