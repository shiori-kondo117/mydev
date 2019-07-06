var url = "jdbc:mariadb://localhost:3306/sample";
var con = java.sql.DriverManager.getConnection(url, 'kondo', 'passwd')
var stmt = con.createStatement()
var rs = stmt.executeQuery('select * from person')
while(rs.next()){
    print(rs.getString('name') + ':' + rs.getInt('age'))
}
print('ok')
