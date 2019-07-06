var url = "jdbc:mariadb://192.168.1.103:3306/sample"
var con = java.sql.DriverManager.getConnection(url, 'admin', 'P@ssw0rd')
var stmt = con.createStatement()
var rs = stmt.executeQuery('select * from test')
while(rs.next()){
    print(rs.getInt('id') + ':' + rs.getString('item'))
}
print('ok')
