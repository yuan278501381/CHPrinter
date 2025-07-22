
Imports System.Data.SqlClient

Module Modue_Global

    Friend ConnectionString As String = "" '连接字符串

    Friend oLayoutManager As New LayoutManager

    '返回一个内存数据表
    Friend Function Rsco(ByVal SqlString As String, Optional conn_str As String = "") As System.Data.DataTable

        If conn_str = "" Then conn_str = ConnectionString

        Using conn As New SqlConnection(conn_str)
            Using cmd As New SqlCommand(SqlString, conn)

                conn.Open()

                Dim Adapter As New System.Data.SqlClient.SqlDataAdapter
                Dim Dt As New System.Data.DataTable
                Adapter.SelectCommand = cmd
                Adapter.Fill(Dt)
                Return Dt

            End Using
        End Using

    End Function


    '执行一段SQL查询，并返回影响条数
    Friend Function DoSql(ByVal SqlString As String, Optional conn_str As String = "") As Integer

        If conn_str = "" Then conn_str = ConnectionString

        Using conn As New SqlConnection(conn_str)
            Using cmd As New SqlCommand(SqlString, conn)
                cmd.CommandText = SqlString
                conn.Open()
                Return cmd.ExecuteNonQuery
            End Using
        End Using

    End Function


    ''' <summary>
    ''' 获得一个连接字符串
    ''' </summary>
    ''' <param name="oSysSt"></param>
    ''' <returns></returns>
    Friend Function GetConnectString(ByRef oSysSt As SysSetting) As String

        Try
            Dim str As String = ""

            If oSysSt.SQL_Server = "" Then Throw New Exception("服务器地址")
            If oSysSt.SQL_DBName = "" Then Throw New Exception("数据库名称")

            If oSysSt.SQL_User = "" Then Throw New Exception("数据库登陆用户名")
            If oSysSt.SQL_Psw = "" Then Throw New Exception("数据库登录密码")
            str = "Persist Security Info=False;User ID=" & oSysSt.SQL_User & ";Password=" & oSysSt.SQL_Psw & ";Initial Catalog=" _
                      & oSysSt.SQL_DBName & ";Server=" & oSysSt.SQL_Server

            Return str

        Catch ex As Exception
            Dim msg As String = "未能获取连接信息：" & ex.Message & ",请到系统设置中补充完整，然后再次尝试！"
            Throw New Exception(msg)
        End Try

    End Function


End Module
