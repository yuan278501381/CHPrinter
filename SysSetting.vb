
Imports System.ComponentModel
Imports System.IO
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Security
Imports System.Security.Cryptography
Imports System.Xml.Linq


'系统设置类
Public Class SysSetting

#Region "静态字段"

    Private configFilePath As String = ""    '配置文件所在路径和文件名


    Public defaultPrinter As String = ""     '上次选择的默认打印机
    Public defaultPrinterFlag As String = ""     '上次选择的默认打印机标识


    Public tb_printers As String = "@CH_LBL_PRINTER"         '存放打印机的表名
    Public fld_printerCode As String = "Code"
    Public fld_printerName As String = "Name"


    Public SQL_Server As String = ""         'SQL服务器的地址
    Public SQL_DBName As String = ""            'SQL账套名
    Public SQL_User As String = "sa"         'SQL的用户
    Public SQL_Psw As String = ""             'SQL的密码

    Public tb_Data As String = "@CH_STGRPT_IVL"             '打印信息表
    Public fld_data_id As String = "TransSeq"                '打印信息表的主键：唯一自增
    Public fld_data_printerFlag As String = "Printer"       '打印信息表记录打印机标识的字段
    Public fld_data_status As String = "Printed"            '打印信息表的状态字段
    Public fld_data_date As String = "PrintDate"            '打印信息表的实际打印日期字段
    Public fld_data_time As String = "PrintTime"            '打印信息表的实际打印时间字段
    Public fld_data_template As String = "PrintTemplate"    '打印信息表记录打印模板标识的字段


    Public tb_tmplt As String = "@CH_LBL_SETTING"            '标签模板表 - 主表
    Public tb_tmplt_child As String = "@CH_LBL_SETTING_1"    '标签模板表 - 子表
    Public fld_tmpl_code As String = "Code"                  '存放标签模板代码的字段
    Public fld_tmpl_fileName As String = "U_FileName"        '存放标签模板文件名的字段
    Public fld_tmpl_url As String = "U_Url"                  '存放标签模板文件路径的字段
    Public fld_tmpl_source As String = "U_Source"            '在标签模板子表中，存放对应关系-源打印字段的字段名
    Public fld_tmpl_target As String = "U_Target"            '在标签模板子表中，存放对应关系-标签位置的字段名

#End Region

#Region "检查定义完整性"


    '是否定义完整，可以监听了
    Sub IsReadyListen()
        If (tb_Data = "") Then Throw New Exception("未定义打印信息表的表名"）
        If (fld_data_printerFlag = "") Then Throw New Exception("未定义打印信息表记录打印机标识的字段"）
        If (fld_data_status = "") Then Throw New Exception("未定义打印信息表的状态字段"）
        If (fld_data_date = "") Then Throw New Exception("未定义打印信息表的实际打印日期字段"）
        If (fld_data_time = "") Then Throw New Exception("未定义打印信息表的实际打印时间字段"）
        If (fld_data_template = "") Then Throw New Exception("未定义打印信息表记录打印模板标识的字段"）

    End Sub


#End Region

#Region "存取配置文件"

    '从本地配置文件中读入设置，如果文件不存在，则新建一个
    Sub LoadFromFile()

        ' 获取当前执行的exe文件的目录
        Dim exePath As String = Assembly.GetExecutingAssembly().Location
        Dim exeDirectory As String = Path.GetDirectoryName(exePath)

        ' 定义配置文件所在的子目录和文件名
        Dim configDirectory As String = Path.Combine(exeDirectory, "config")
        Dim configFileName As String = "config.ini"
        configFilePath = Path.Combine(configDirectory, configFileName)

        ' 检查配置文件所在的目录是否存在，如果不存在则创建
        If Not Directory.Exists(configDirectory) Then
            Directory.CreateDirectory(configDirectory)
        End If

        ' 检查配置文件是否存在
        If Not File.Exists(configFilePath) Then
            ' 如果配置文件不存在，则创建一个新的配置文件并写入XML内容
            SaveToFile()
        End If

        ' 从config.ini文件读取XML内容到XElement
        Try
            Dim oXml As XElement = XElement.Load(configFilePath)
            SQL_Server = oXml.<sqlconnect>(0).<server>(0).Value
            SQL_DBName = oXml.<sqlconnect>(0).<DBName>(0).Value
            SQL_User = oXml.<sqlconnect>(0).<DBUser>(0).Value
            If (oXml.<sqlconnect>(0).<DBPsw>(0).Value <> "") Then
                ' 进行Base64解码
                Dim decodedBytes As Byte() = System.Convert.FromBase64String(oXml.<sqlconnect>(0).<DBPsw>(0).Value)
                ' 将字节数组转换回字符串
                SQL_Psw = System.Text.Encoding.UTF8.GetString(decodedBytes)
            End If

            tb_printers = oXml.<Printers>(0).<TableName>(0).Value
            fld_printerCode = oXml.<Printers>(0).<CodeFld>(0).Value
            fld_printerName = oXml.<Printers>(0).<NameFld>(0).Value

            defaultPrinter = oXml.<dftValue>(0).<defaultPrinter>.Value
            defaultPrinterFlag = oXml.<dftValue>(0).<defaultPrinterFlag>.Value

            tb_Data = oXml.<PrintData>(0).<tbName>.Value
            fld_data_printerFlag = oXml.<PrintData>(0).<fldFlag>.Value
            fld_data_status = oXml.<PrintData>(0).<fldStatus>.Value
            fld_data_date = oXml.<PrintData>(0).<fldDate>.Value
            fld_data_time = oXml.<PrintData>(0).<fldTime>.Value
            fld_data_template = oXml.<PrintData>(0).<fldTmptCode>.Value
            fld_data_id = oXml.<PrintData>(0).<fldId>.Value

            tb_tmplt = oXml.<Template>(0).<tbName>.Value
            tb_tmplt_child = oXml.<Template>(0).<tbChildName>.Value
            fld_tmpl_code = oXml.<Template>(0).<fldCode>.Value
            fld_tmpl_fileName = oXml.<Template>(0).<fldFileName>.Value
            fld_tmpl_url = oXml.<Template>(0).<fldUrl>.Value
            fld_tmpl_source = oXml.<Template>(0).<fldSource>.Value
            fld_tmpl_target = oXml.<Template>(0).<fldTarget>.Value

        Catch ex As Exception
            Throw New Exception("读取配置文件时出现错误:" & ex.Message)
        End Try

    End Sub

    '保存自己到XML本地文件
    Sub SaveToFile()

        Dim oXml As XElement = <configuration>
                                   <dftValue>
                                       <defaultPrinter></defaultPrinter>
                                       <defaultPrinterFlag></defaultPrinterFlag>
                                   </dftValue>
                                   <sqlconnect>
                                       <server></server>
                                       <DBName></DBName>
                                       <DBUser></DBUser>
                                       <DBPsw></DBPsw>
                                   </sqlconnect>
                                   <Printers>
                                       <TableName></TableName>
                                       <CodeFld></CodeFld>
                                       <NameFld></NameFld>
                                   </Printers>
                                   <PrintData>
                                       <tbName></tbName>
                                       <fldId></fldId>
                                       <fldFlag></fldFlag>
                                       <fldStatus></fldStatus>
                                       <fldDate></fldDate>
                                       <fldTime></fldTime>
                                       <fldTmptCode></fldTmptCode>
                                   </PrintData>
                                   <Template>
                                       <tbName></tbName>
                                       <tbChildName></tbChildName>
                                       <fldCode></fldCode>
                                       <fldFileName></fldFileName>
                                       <fldUrl></fldUrl>
                                       <fldSource></fldSource>
                                       <fldTarget></fldTarget>
                                   </Template>
                               </configuration>

        oXml.<sqlconnect>(0).<server>(0).Value = SQL_Server
        oXml.<sqlconnect>(0).<DBName>(0).Value = SQL_DBName
        oXml.<sqlconnect>(0).<DBUser>(0).Value = SQL_User

        If (SQL_Psw <> "") Then
            ' 将字符串转换为字节数组
            Dim bytesToEncode As Byte() = System.Text.Encoding.UTF8.GetBytes(SQL_Psw)
            ' 进行Base64编码
            Dim encodedText As String = System.Convert.ToBase64String(bytesToEncode)

            oXml.<sqlconnect>(0).<DBPsw>(0).Value = encodedText
        End If


        oXml.<Printers>(0).<TableName>(0).Value = tb_printers
        oXml.<Printers>(0).<CodeFld>(0).Value = fld_printerCode
        oXml.<Printers>(0).<NameFld>(0).Value = fld_printerName

        oXml.<dftValue>(0).<defaultPrinter>.Value = defaultPrinter
        oXml.<dftValue>(0).<defaultPrinterFlag>.Value = defaultPrinterFlag

        oXml.<PrintData>(0).<tbName>.Value = tb_Data
        oXml.<PrintData>(0).<fldId>.Value = fld_data_id
        oXml.<PrintData>(0).<fldFlag>.Value = fld_data_printerFlag
        oXml.<PrintData>(0).<fldStatus>.Value = fld_data_status
        oXml.<PrintData>(0).<fldDate>.Value = fld_data_date
        oXml.<PrintData>(0).<fldTime>.Value = fld_data_time
        oXml.<PrintData>(0).<fldTmptCode>.Value = fld_data_template

        oXml.<Template>(0).<tbName>.Value = tb_tmplt
        oXml.<Template>(0).<tbChildName>.Value = tb_tmplt_child
        oXml.<Template>(0).<fldCode>.Value = fld_tmpl_code
        oXml.<Template>(0).<fldFileName>.Value = fld_tmpl_fileName
        oXml.<Template>(0).<fldUrl>.Value = fld_tmpl_url
        oXml.<Template>(0).<fldSource>.Value = fld_tmpl_source
        oXml.<Template>(0).<fldTarget>.Value = fld_tmpl_target

        ' 保存XML内容到config.ini文件
        oXml.Save(configFilePath)


    End Sub

#End Region

#Region "属性定义"

#Region "SQL连接信息"

    <Category("SQL 连接信息"), Description("服务器地址")>
    Public Property DataSource As String
        Get
            Return SQL_Server
        End Get
        Set(value As String)
            SQL_Server = value
        End Set
    End Property

    <Category("SQL 连接信息"), Description("数据库名称")>
    Public Property DBName As String
        Get
            Return SQL_DBName
        End Get
        Set(value As String)
            SQL_DBName = value
        End Set
    End Property

    <Category("SQL 连接信息"), Description("用于连接数据库的用户名")>
    Public Property DBUser As String
        Get
            Return SQL_User
        End Get
        Set(value As String)
            SQL_User = value
        End Set
    End Property


    <Category("SQL 连接信息"), Description("用于数据库连接的密码。"), PasswordPropertyText(True)>
    Public Property Password As String
        Get
            Return SQL_Psw
        End Get
        Set(value As String)
            SQL_Psw = value
        End Set
    End Property
#End Region

#Region "打印机标识数据源"

    <Category("打印机标识数据源"), Description("存放打印机标识清单的表名")>
    Public Property PrinterTableName As String
        Get
            Return tb_printers
        End Get
        Set(value As String)
            tb_printers = value
        End Set
    End Property

    <Category("打印机标识数据源"), Description("存放打印机标识的字段名")>
    Public Property PrinterCodeFld As String
        Get
            Return fld_printerCode
        End Get
        Set(value As String)
            fld_printerCode = value
        End Set
    End Property


    <Category("打印机标识数据源"), Description("存放打印机描述的字段名")>
    Public Property PrinterNameFld As String
        Get
            Return fld_printerName
        End Get
        Set(value As String)
            fld_printerName = value
        End Set
    End Property


#End Region

#Region "打印内容来源"

    <Category("打印内容来源"), Description("存放打印内容的SQL表名")>
    Public Property TableNameOfData As String
        Get
            Return tb_Data
        End Get
        Set(value As String)
            tb_Data = value
        End Set
    End Property

    <Category("打印内容来源"), Description("在打印内容的表的唯一标识主键字段名")>
    Public Property FieldNameOfID As String
        Get
            Return fld_data_id
        End Get
        Set(value As String)
            fld_data_id = value
        End Set
    End Property


    <Category("打印内容来源"), Description("在打印内容的表中，记录打印机标识的字段")>
    Public Property FieldNameOfPrinterFlag As String
        Get
            Return fld_data_printerFlag
        End Get
        Set(value As String)
            fld_data_printerFlag = value
        End Set
    End Property

    <Category("打印内容来源"), Description("在打印内容的表中，记录打印状态的字段")>
    Public Property FieldNameOfStatus As String
        Get
            Return fld_data_status
        End Get
        Set(value As String)
            fld_data_status = value
        End Set
    End Property

    <Category("打印内容来源"), Description("在打印内容的表中，记录实际打印日期的字段")>
    Public Property FieldNameOfDate As String
        Get
            Return fld_data_date
        End Get
        Set(value As String)
            fld_data_date = value
        End Set
    End Property

    <Category("打印内容来源"), Description("在打印内容的表中，记录实际打印时间的字段")>
    Public Property FieldNameOfTime As String
        Get
            Return fld_data_time
        End Get
        Set(value As String)
            fld_data_time = value
        End Set
    End Property

    <Category("打印内容来源"), Description("在打印内容的表中，记录打印模板标识的字段")>
    Public Property FieldNameOfTemplateCode As String
        Get
            Return fld_data_template
        End Get
        Set(value As String)
            fld_data_template = value
        End Set
    End Property

#End Region

#Region "标签模板表 - 主表"

    <Category("标签模板表 - 主表"), Description("SQL表名")>
    Public Property TableNameOfTemplate As String
        Get
            Return tb_tmplt
        End Get
        Set(value As String)
            tb_tmplt = value
        End Set
    End Property

    <Category("标签模板表 - 主表"), Description("存放标签模板代码的字段")>
    Public Property FieldNameOfCode As String
        Get
            Return fld_tmpl_code
        End Get
        Set(value As String)
            fld_tmpl_code = value
        End Set
    End Property

    <Category("标签模板表 - 主表"), Description("存放标签模板文件名的字段")>
    Public Property FieldNameOfFileName As String
        Get
            Return fld_tmpl_fileName
        End Get
        Set(value As String)
            fld_tmpl_fileName = value
        End Set
    End Property


    <Category("标签模板表 - 主表"), Description("存放标签模板文件路径的字段")>
    Public Property FieldNameOfPath As String
        Get
            Return fld_tmpl_url
        End Get
        Set(value As String)
            fld_tmpl_url = value
        End Set
    End Property

#End Region


#Region "标签模板表 - 子表"

    <Category("标签模板表 - 子表"), Description("SQL表名")>
    Public Property TableNameOfTemplateChild As String
        Get
            Return tb_tmplt_child
        End Get
        Set(value As String)
            tb_tmplt_child = value
        End Set
    End Property

    <Category("标签模板表 - 子表"), Description("在标签模板子表中，存放对应关系-源打印字段的字段名")>
    Public Property FieldNameOfSource As String
        Get
            Return fld_tmpl_source
        End Get
        Set(value As String)
            fld_tmpl_source = value
        End Set
    End Property

    <Category("标签模板表 - 子表"), Description("在标签模板子表中，存放对应关系-标签位置的字段名")>
    Public Property FieldNameOfTarget As String
        Get
            Return fld_tmpl_target
        End Get
        Set(value As String)
            fld_tmpl_target = value
        End Set
    End Property

#End Region

#End Region

End Class
