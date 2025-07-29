Imports System.Threading
Imports AlibreAddOn
Imports AlibreX
Namespace alibretestbedaddonForAlibreDesign
    Public Class MenuItem
        Public Property Id As Integer
        Public Property Text As String
        Public Property ToolTip As String
        Public Property Icon As String
        Public Property Command As Func(Of IADSession, IAlibreAddOnCommand)
        Public Property SubItems As List(Of MenuItem)
        Public Sub New(id As Integer, text As String, Optional toolTip As String = "", Optional icon As String = "", Optional command As Func(Of IADSession, IAlibreAddOnCommand) = Nothing)
            Me.Id = id
            Me.Text = text
            Me.ToolTip = toolTip
            Me.Icon = icon
            Me.Command = command
            Me.SubItems = New List(Of MenuItem)()
        End Sub
        Public Sub AddSubItem(subItem As MenuItem)
            SubItems.Add(subItem)
        End Sub
        Public Function DummyFunction(session As IADSession) As IAlibreAddOnCommand
            MsgBox("{CLICK-TEST}")
            Return Nothing
        End Function
    End Class
End Namespace