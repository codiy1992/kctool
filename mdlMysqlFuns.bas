Attribute VB_Name = "mdlMysqlFuns"
    ' 定义并创建数据库连接和访问对象
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
      
    ' 定义数据库连接字符串变量
    Dim strCn As String
      
    ' 定义数据库连接参数变量
    Dim db_host As String
    Dim db_user As String
    Dim db_pass As String
    Dim db_data As String
      
    ' 定义 SQL 语句变量
    Dim sql As String
      
    ' 初始化数据库连接变量
    db_host = "localhost"
    db_user = "root"
    db_pass = "root"
    db_data = "test"
      
    ' MySQL ODBC 连接参数
    '+------------+---------------------+----------------------------------+
    '| 参数名     | 默认值              | 说明                             |
    '+------------+------------------------------------------------------–+
    '| user       | ODBC (on Windows)   | MySQL 用户名                     |
    '| server     | localhost           | MySQL 服务器地址                 |
    '| database   |                     | 默认连接数据库                   |
    '| option     | 0                   | 参数用以指定连接的工作方式       |
    '| port       | 3306                | 连接端口                         |
    '| stmt       |                     | 一段声明, 可以在连接数据库后运行 |
    '| password   |                     | MySQL 用户密码                   |
    '| socket     |                     | (略)                             |
    '+------------+---------------------+----------------------------------+
      
    ' 详细查看官方说明
    ' http://dev.mysql.com/doc/refman/5.0/en/myodbc-configuration-connection-parameters.html
      
    strCn = "DRIVER={MySQL ODBC 5.3 Driver};" & _
             "SERVER=" & db_host & ";" & _
             "DATABASE=" & db_data & ";" & _
             "UID=" & db_user & ";PWD=" & db_pass & ";" & _
             "OPTION=3;stmt=SET NAMES GB2312"
      
    ' stmt=SET NAMES GB2312
    ' 这句是设置数据库编码方式
    ' 中文操作系统需要设置成 GB2312
    ' 这样中文才不会有问题
    ' 版本要求 mysql 4.1+
      
    ' 连接数据库
    cn.Open strCn
    ' 设置该属性, 使 recordcount 和 absolutepage 属性可用
    cn.CursorLocation = adUseClient
      
    ' 访问表users
    sql = "select 102"
    rs.Open sql, cn
    MsgBox rs.RecordCount

