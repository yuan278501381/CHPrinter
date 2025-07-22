
''' <summary>
''' 打印标签的管理类
''' </summary>
Friend Class LayoutManager

    Private layouts As New Hashtable            '存放每一个标签信息



    Friend Function GetLayout(ByVal code As String, ByRef obj As SysSetting) As LayoutTemlate

        If layouts.ContainsKey(code) Then Return CType(layouts.Item(code), LayoutTemlate)

        '没有保存，获取新的
        Dim rlt As New LayoutTemlate
        rlt.Code = code

        '主表信息
        Dim sb As New Text.StringBuilder("select * from [")
        sb.Append(obj.tb_tmplt).Append("] where [").Append(obj.fld_tmpl_code).Append("] =N'").Append(code).Append("'")

        Dim dt As Data.DataTable = Rsco(sb.ToString)
        If dt.Rows.Count = 0 Then Throw New Exception("错误：没有找到标签模板定义，代码：" + code)

        rlt.FileName = dt.Rows.Item(0).Item(obj.fld_tmpl_fileName)
        rlt.PathUrl = dt.Rows.Item(0).Item(obj.fld_tmpl_url)

        '子表信息
        sb.Clear()
        sb.Append("select * from [")
        sb.Append(obj.tb_tmplt_child).Append("] where [").Append(obj.fld_tmpl_code).Append("] =N'").Append(code).Append("'")

        dt = Rsco(sb.ToString)
        For i As Integer = 0 To dt.Rows.Count - 1
            Dim src As String = dt.Rows.Item(i).Item(obj.fld_tmpl_source).ToString
            Dim tag As String = dt.Rows.Item(i).Item(obj.fld_tmpl_target).ToString
            If Not (String.IsNullOrWhiteSpace(src) Or String.IsNullOrWhiteSpace(tag)) Then
                rlt.FldMappingsList.Add(Tuple.Create(src, tag))
            End If

        Next

        layouts.Add(code, rlt)
        Return rlt
    End Function


End Class
