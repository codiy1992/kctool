Attribute VB_Name = "mdlMysqlFuns"
    ' ���岢�������ݿ����Ӻͷ��ʶ���
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
      
    ' �������ݿ������ַ�������
    Dim strCn As String
      
    ' �������ݿ����Ӳ�������
    Dim db_host As String
    Dim db_user As String
    Dim db_pass As String
    Dim db_data As String
      
    ' ���� SQL ������
    Dim sql As String
      
    ' ��ʼ�����ݿ����ӱ���
    db_host = "localhost"
    db_user = "root"
    db_pass = "root"
    db_data = "test"
      
    ' MySQL ODBC ���Ӳ���
    '+------------+---------------------+----------------------------------+
    '| ������     | Ĭ��ֵ              | ˵��                             |
    '+------------+------------------------------------------------------�C+
    '| user       | ODBC (on Windows)   | MySQL �û���                     |
    '| server     | localhost           | MySQL ��������ַ                 |
    '| database   |                     | Ĭ���������ݿ�                   |
    '| option     | 0                   | ��������ָ�����ӵĹ�����ʽ       |
    '| port       | 3306                | ���Ӷ˿�                         |
    '| stmt       |                     | һ������, �������������ݿ������ |
    '| password   |                     | MySQL �û�����                   |
    '| socket     |                     | (��)                             |
    '+------------+---------------------+----------------------------------+
      
    ' ��ϸ�鿴�ٷ�˵��
    ' http://dev.mysql.com/doc/refman/5.0/en/myodbc-configuration-connection-parameters.html
      
    strCn = "DRIVER={MySQL ODBC 5.3 Driver};" & _
             "SERVER=" & db_host & ";" & _
             "DATABASE=" & db_data & ";" & _
             "UID=" & db_user & ";PWD=" & db_pass & ";" & _
             "OPTION=3;stmt=SET NAMES GB2312"
      
    ' stmt=SET NAMES GB2312
    ' ������������ݿ���뷽ʽ
    ' ���Ĳ���ϵͳ��Ҫ���ó� GB2312
    ' �������ĲŲ���������
    ' �汾Ҫ�� mysql 4.1+
      
    ' �������ݿ�
    cn.Open strCn
    ' ���ø�����, ʹ recordcount �� absolutepage ���Կ���
    cn.CursorLocation = adUseClient
      
    ' ���ʱ�users
    sql = "select 102"
    rs.Open sql, cn
    MsgBox rs.RecordCount

