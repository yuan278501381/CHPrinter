


Imports System.Runtime



Public Class FSysSt

    ' 定义事件
    Public Event AfterFormClosed()

    ' 在关闭时触发事件
    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        MyBase.OnFormClosing(e)

        RaiseEvent AfterFormClosed()
    End Sub

    Friend Sub SetObject(ByRef st As SysSetting)
        PropertyGrid1.SelectedObject = st
    End Sub

End Class