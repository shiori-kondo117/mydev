var list = java.util.Arrays.asList(1, 2, 3, 4, 5, 6, 7, 8);
//var odd = list.stream().filter(
//  function(i) {
//  return i % 2 == 0;
//});
var odd = list.stream().filter(
  function(i) i % 2 == 0);
odd.forEach(function(i) {
print(">>> " + i);
});
