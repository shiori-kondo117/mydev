var prop = new java.util.Properties();
prop.setProperty('user', 'kondo');
prop.setProperty('password', 'passwd');
var url = "jdbc:mariadb://localhost:3306/sample";
var driver = new org.mariadb.jdbc.Driver();
var con = driver.connect(url, prop);
var stmt = con.createStatement();
var rs = stmt.executeQuery('select * from person');
while(rs.next()){
    print(rs.getString('name') + ':' + rs.getInt('age'));
}
print('ok');
