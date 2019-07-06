var stack = 
new java.util.LinkedList();
[1, 2, 3, 4].forEach(function(item) {
  stack.push(item);
});
var sorted = stack
	.stream()
	.sorted()
	.toArray();
print(sorted);

var jsArray = java.from(sorted);
print(jsArray);

var javaArray = 
java.to(jsArray);
print(javaArray);
