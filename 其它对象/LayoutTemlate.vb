Imports Seagull.BarTender.Print

Imports System.IO

''' <summary>
''' 打印模板定义（使用 BarTender 2022 .NET SDK）
''' </summary>
Friend Class LayoutTemlate

#Region "字段"

    Friend Code As String = ""
    Friend FileName As String = ""
    Friend PathUrl As String = ""
    Friend FldMappingsList As New List(Of Tuple(Of String, String))

#End Region

#Region "动作"

    ''' <summary>
    ''' 打印指定的记录
    ''' </summary>
    Friend Sub Print(ByRef row As Data.DataRow, ByVal Printer As String, ByRef obj As SysSetting)
        ' 获取记录ID用于日志
        Dim recordId As String = row.Item(obj.fld_data_id).ToString()
        Dim logMessage As Action(Of String) = AddressOf MainForm.Instance.AddLogo

        Try
            ' 1. 检查模板文件是否存在
            Dim tmpFile As String = Path.Combine(PathUrl, FileName)
            logMessage($"【{recordId}】开始打印处理，使用模板: {Code}")

            If Not File.Exists(tmpFile) Then
                logMessage($"❌ 错误：模板文件不存在: {tmpFile}")
                Throw New Exception($"模板文件不存在: {tmpFile}")
            End If
            logMessage($"✅ 模板文件验证成功: {tmpFile}")

            ' 2. 创建打印引擎
            Using btEngine As New Engine(True)
                logMessage($"【{recordId}】正在打开模板文件...")
                Dim format As LabelFormatDocument = btEngine.Documents.Open(tmpFile)

                ' 3. 设置打印机
                If Not String.IsNullOrEmpty(Printer) Then
                    format.PrintSetup.PrinterName = Printer
                    logMessage($"🖨️ 设置打印机: {Printer}")
                Else
                    logMessage("⚠️ 使用默认打印机")
                End If

                ' 4. 设置字段值
                logMessage($"【{recordId}】正在设置字段值...")
                For Each mapping In FldMappingsList
                    Try
                        Dim fieldValue = row.Item(mapping.Item1).ToString()
                        format.SubStrings(mapping.Item2).Value = fieldValue
                        logMessage($"  字段映射: [{mapping.Item1}] → [{mapping.Item2}] = '{fieldValue}'")
                    Catch ex As Exception
                        logMessage($"❌ 字段设置失败: {mapping.Item1} → {mapping.Item2}")
                        Throw New Exception($"字段设置失败: {mapping.Item1} → {mapping.Item2}", ex)
                    End Try
                Next

                ' 5. 执行打印
                logMessage($"【{recordId}】正在发送打印任务...")
                Dim jobName As String = $"autoPrint_{recordId}"
                Dim result As Result = format.Print(jobName, 1)

                If result = Result.Failure Then
                    logMessage($"❌ 打印失败! BarTender 状态: {result}")
                    Throw New Exception($"BarTender 打印失败: {result}")
                End If
                logMessage($"✅ 打印任务发送成功! 作业名称: {jobName}")

                ' 6. 关闭格式文档
                format.Close(SaveOptions.DoNotSaveChanges)
            End Using

            ' 7. 更新数据库状态
            logMessage($"【{recordId}】正在更新打印状态...")
            Dim sb As New Text.StringBuilder("update [")
            sb.Append(obj.tb_Data).Append("] set [").Append(obj.fld_data_status).Append("] = 'Y',[")
            sb.Append(obj.fld_data_date).Append("] = convert(date,'").Append(Now.ToString("yyyyMMdd")).Append("',112),[")
            sb.Append(obj.fld_data_time).Append("] = ").Append(Now.ToString("HHmm")).Append(" where [")
            sb.Append(obj.fld_data_id).Append("] = ").Append(recordId)

            Try
                DoSql(sb.ToString)
                logMessage($"✅ 记录 {recordId} 状态更新成功")
            Catch sqlEx As Exception
                logMessage($"❌ 状态更新失败! SQL: {sb.ToString()}")
                Throw New Exception($"更新打印状态失败: {sqlEx.Message}", sqlEx)
            End Try

        Catch ex As Exception
            logMessage($"❌【{recordId}】打印处理失败: {ex.Message}")
            Throw  ' 重新抛出异常
        Finally
            logMessage($"【{recordId}】打印流程结束")
        End Try
    End Sub

#End Region
End Class