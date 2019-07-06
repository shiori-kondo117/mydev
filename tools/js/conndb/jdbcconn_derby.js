var url = "jdbc:derby://localhost:1527//C:/Users/oyaji/Documents/mydata/db/derby/DerbyTestDB"
var con = java.sql.DriverManager.getConnection(url)
var stmt = con.createStatement()
var rs = stmt.executeQuery('select * from member')
while(rs.next()){
    print(rs.getInt('id') + ',' + rs.getString('name'))
}
rs.close();
stmt.close();
con.close();

print('ok')
