var CollectionsAndFiles = new JavaImporter(
    java.util,
    java.io,
    java.nio);
 
with (CollectionsAndFiles) {
  var files = new LinkedHashSet();
  files.add(new File("Plop"));
  files.add(new File("Foo"));
  files.add(new File("w00t.js"));
}
