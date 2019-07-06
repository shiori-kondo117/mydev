var LinkedList = Java.type(
"java.util.LinkedList");
var primitiveInt = Java.type(
"int");
var arrayOfInts = Java.type(
"int[]");
var list = new LinkedList;
list.add(1);
list.add(2);
print(LinkedList.class);
print(list.getClass().static);
print(LinkedList.class === list.getClass());
print(list.getClass().static === LinkedList);
