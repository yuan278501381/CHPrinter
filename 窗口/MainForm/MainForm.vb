
Imports System
Imports System.Xml
Imports System.Security.Cryptography
Imports System.Data.SqlClient
Imports System.Drawing.Printing
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Imports System.Linq
Imports System.Windows
Imports System.IO
Imports System.Reflection

Public Class MainForm

    Public Shared Instance As MainForm

    Private _object As SysSetting   '数据源

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Instance = Me

        Try
            '===========================================================
            _object = New SysSetting()  '初始化数据源
            _object.LoadFromFile()
            '===========================================================

            RefreshFormUI()

        Catch ex As Exception
            MsgBox("出现意外错误，请联系管理员：" & ex.Message)
            Me.Close()
            End
        End Try



    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If MessageBox.Show("你确定要清空所有的日志吗？", "确定", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            txtLogo.Text = ""
        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim oFSysSt As New FSysSt
        oFSysSt.SetObject(Me._object)
        AddHandler oFSysSt.AfterFormClosed, AddressOf RefreshFormUI
        oFSysSt.ShowDialog()

    End Sub


    Private Sub ShowMessageOnStatusStrip(txt As String)
        ' 显示消息
        error_info.Text = txt
        error_info.BackColor = Color.Red
        error_info.ForeColor = Color.White

        AddLogo(txt)

        ' 启动Timer
        Timer2.Enabled = True

    End Sub


    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        ' 清除文本
        error_info.Text = ""

        error_info.BackColor = Color.Transparent
        error_info.ForeColor = Color.Black

        ' 停止Timer
        Timer2.Enabled = False
    End Sub


    '读取配置文件后，刷新界面
    Private Sub RefreshFormUI()

        Try
            '===========================================================
            ' 获取并列出所有安装的打印机
            ComboBox1.Items.Clear()
            For Each printer As String In PrinterSettings.InstalledPrinters
                ComboBox1.Items.Add(printer)
            Next


            ' 查找项列表中是否存在指定的默认打印机
            Dim itemIndex As Integer = ComboBox1.FindStringExact(_object.defaultPrinter)

            ' 如果存在，则设置ComboBox以显示该文本
            If itemIndex >= 0 Then
                ComboBox1.SelectedIndex = itemIndex
            Else
                ' 如果不存在，则显示空白
                ComboBox1.Text = ""  ' 清除当前显示的文本
                ComboBox1.SelectedIndex = -1 ' 确保没有选中任何项
            End If
            '===========================================================


        Catch ex As Exception
            MsgBox("出现意外错误，请联系管理员：" & ex.Message)
            Me.Close()
            End
        End Try


        '===========================================================
        '连接数据库 - 并获得打印机标认清单
        Try
            Dim sqlstring = GetConnectString(_object) '获得一个连接字符串

            Try
                '获得打印机标识清单
                Dim sb As New Text.StringBuilder("select [")
                sb.Append(_object.fld_printerCode).Append("],[")
                sb.Append(_object.fld_printerName).Append("] from [")
                sb.Append(_object.tb_printers).Append("] ")

                Dim dt As DataTable = Rsco(sb.ToString, sqlstring)

                Dim options As New List(Of KeyValuePair(Of String, String))
                For i As Integer = 0 To dt.Rows.Count - 1
                    options.Add(New KeyValuePair(Of String, String)(dt.Rows.Item(i).Item(0).ToString, dt.Rows.Item(i).Item(0).ToString & " - " & dt.Rows.Item(i).Item(1).ToString))
                Next

                ' 绑定到ComboBox
                With flagCombox
                    .DataSource = options
                    .DisplayMember = "Value"
                    .ValueMember = "Key"
                End With

                '设置默认值 
                If (_object.defaultPrinterFlag <> "") Then
                    If options.Any(Function(pair) pair.Key = _object.defaultPrinterFlag) Then
                        flagCombox.SelectedValue = _object.defaultPrinterFlag
                    End If
                End If

                ConnectionString = sqlstring '保存连接信息

                flagCombox.Enabled = True '打印机标识下拉框可用

            Catch ex As Exception
                ConnectionString = ""   '清空连接字符串
                ShowMessageOnStatusStrip("错误：未能获得打印标识信息，请检查系统设置！详细信息：" & ex.Message)
            End Try



        Catch ex As Exception
            ShowMessageOnStatusStrip(”错误：“ & ex.Message)
        End Try
        '===========================================================

        _object.SaveToFile() '保存配置

        '===========================================================
        '根据连接结果，让运行按钮和打印机标识下拉框是否可用
        Dim isConnected As Boolean = (ConnectionString <> "")
        flagCombox.Enabled = isConnected
        btnRun.Enabled = isConnected
        If isConnected Then
            StatusLabel.Text = "已连接至数据库：" & _object.DBName
            StatusLabel.ForeColor = Color.Blue

            AddLogo(StatusLabel.Text)
        Else
            StatusLabel.Text = "未连接"
            StatusLabel.ForeColor = Color.Black
        End If
        '===========================================================

    End Sub



    Friend Sub AddLogo(msg As String)
        ' 获取当前日期和时间
        Dim now As DateTime = DateTime.Now

        ' 转换为字符串，格式为“年-月-日 时:分:秒”
        Dim dateTimeString As String = now.ToString("yyyy-MM-dd HH:mm:ss")
        Dim sb As New System.Text.StringBuilder(dateTimeString)
        sb.Append(vbTab).Append(msg).Append(vbCrLf).Append(txtLogo.Text)

        If txtLogo.InvokeRequired Then
            txtLogo.Invoke(New Action(Sub() txtLogo.Text = sb.ToString))
        Else
            txtLogo.Text = sb.ToString
        End If
    End Sub

    '保存日志
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Try
            ' 获取当前执行的exe文件的目录
            Dim exePath As String = Assembly.GetExecutingAssembly().Location
            Dim exeDirectory As String = Path.GetDirectoryName(exePath)

            ' 定义日志文件所在的子目录和文件名
            Dim logDirectory As String = Path.Combine(exeDirectory, "log")
            Dim FileName As String = Now.ToString("yyyy-MM-dd_HH_mm_ss")
            FileName = FileName & "_CHPrinter_Run_log.txt"

            FileName = Path.Combine(logDirectory, FileName)

            ' 检查配置文件所在的目录是否存在，如果不存在则创建
            If Not Directory.Exists(logDirectory) Then
                Directory.CreateDirectory(logDirectory)
            End If

            File.WriteAllText(FileName, txtLogo.Text)

            If MessageBox.Show("日志文件已创建到：" & FileName & vbCrLf & vbCrLf & "你需要清空当前窗口中的日志吗？", "清空日志", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                txtLogo.Text = ""
            End If

        Catch ex As Exception
            ShowMessageOnStatusStrip(ex.Message)
        End Try
    End Sub

    Private Sub btnRun_Click(sender As Object, e As EventArgs) Handles btnRun.Click
        '开始运行
        Try
            '======================================================
            '检查和保存默认值 
            Dim printer As String = ComboBox1.Text
            Dim printerFlag As String = flagCombox.SelectedValue.ToString

            If (printer = "") Then Throw New Exception("错误：未选择本地打印机，无法启动！")
            If (printerFlag = "") Then Throw New Exception("错误：未选择打印机标识，无法启动！")

            _object.defaultPrinterFlag = printerFlag
            _object.defaultPrinter = printer
            '======================================================




            '======================================================
            '构建SQL语句 - 只监控未打印的，并且打印标识与自己相同的记录
            Dim sb As New Text.StringBuilder
            sb.Append(" where [")
            sb.Append(_object.fld_data_status).Append("] = 'N' and [")
            sb.Append(_object.fld_data_printerFlag).Append("] = N'").Append(printerFlag).Append("'")

            query_condition = sb.ToString
            query_str = "select [" + _object.fld_data_id + "] from dbo.[" + _object.tb_Data + "] " + query_condition
            '======================================================


            'MsgBox(query_str)
            '======================================================
            ' 开始侦听数据库通知
            AddLogo("开始执行 - 向服务器请求建立连接")

            DoJob() '先运行一次

            StartListening(ConnectionString)
            btnRun.Enabled = False '按钮不可用
            '======================================================

        Catch ex As Exception
            ShowMessageOnStatusStrip(ex.Message)
        End Try
    End Sub


#Region “查询通知相关代码"

    Private query_str As String = ""     '监控打印信息表的SQL查询代码
    Private query_condition As String = ""   '监控/获取打印信息表的条件

    Private Sub StartListening(connectionString As String)
        ' 开启SqlDependency与SQL Server的通信
        SqlDependency.Start(connectionString)
        Try
            Using connection As New SqlConnection(connectionString)
                connection.Open()

                Using command As SqlCommand = connection.CreateCommand()
                    command.CommandText = query_str

                    ' 创建SqlDependency并绑定OnChange事件
                    Dim dependency As New SqlDependency(command)
                    AddHandler dependency.OnChange, AddressOf OnDependencyChange

                    command.ExecuteReader().Dispose()

                    AddLogo("成功与服务器建立连接")
                    sts_run.Text = "正在运行中"

                End Using
            End Using
        Catch ex As Exception
            ShowMessageOnStatusStrip("发生错误(StartListening): " & ex.Message)
            SqlDependency.Stop(connectionString)
            sts_run.Text = "未运行"
        End Try
    End Sub



    Private Sub OnDependencyChange(sender As Object, e As SqlNotificationEventArgs)

        '=====================================================================
        ' 移除事件处理程序，以便垃圾回收器可以回收依赖对象
        Dim dependency As SqlDependency = CType(sender, SqlDependency)
        RemoveHandler dependency.OnChange, AddressOf OnDependencyChange
        '=====================================================================

        StartListening(ConnectionString)

        DoJob()



    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        End
    End Sub

    Private Sub DoJob()
        Try
            '=====================================================================
            Dim printer As String = CType(Me.Invoke(New Func(Of String)(Function() ComboBox1.Text)), String)
            Dim printerFlag As String = CType(Me.Invoke(New Func(Of String)(Function() Convert.ToString(flagCombox.SelectedValue))), String)

            ' 在这里处理通知事件
            AddLogo("获得需要打印的任务，正在启动处理,打印机标识：" + printerFlag + " 打印机：" + printer)

            '先把变更的数据读入到本地
            Dim query_data As String = "select * from dbo.[" + _object.tb_Data + "] " + query_condition
            Dim dt_data As System.Data.DataTable = Rsco(query_data)


            '最后对读入到本地的数据进行处理
            Dim ct As Integer = dt_data.Rows.Count
            AddLogo("处理需要打印的数据，共计：" + ct.ToString + " 条记录")

            For i As Integer = 0 To ct - 1

                Dim flag As String = (i + 1).ToString + "/" + ct.ToString

                AddLogo(flag + "正在获得打印模板的详细定义...")
                '==================================================================
                Dim tmplateCode As String = dt_data.Rows.Item(i).Item(_object.fld_data_template)
                Dim layout As LayoutTemlate = oLayoutManager.GetLayout(tmplateCode, _object)

                AddLogo(flag + "正在打印到：" + printer)
                layout.Print(dt_data.Rows.Item(i), printer, _object)
                '==================================================================

            Next


            '=====================================================================
        Catch ex As Exception
            If Me.InvokeRequired Then
                Me.Invoke(New Action(Sub() ShowMessageOnStatusStrip("处理打印任务时发生错误: " & ex.Message)))
            Else
                ShowMessageOnStatusStrip("处理打印任务时发生错误: " & ex.Message)
            End If


        End Try
    End Sub

    Private Sub Label26_Click(sender As Object, e As EventArgs) Handles Label26.Click

    End Sub

#End Region
End Class
