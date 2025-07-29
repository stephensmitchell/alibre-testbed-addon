Imports AlibreX
Imports AlibreAddOn
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Public Class AlibreObjectViewerCommand
    Implements IAlibreAddOnCommand
    Private session As IADSession
    Public Property CommandSite As IADAddOnCommandSite Implements IAlibreAddOnCommand.CommandSite
    Public Sub OnComplete() Implements IAlibreAddOnCommand.OnComplete
        Try
            Dim hook As IAutomationHook = CType(Marshal.GetActiveObject("AlibreX.AutomationHook"), IAutomationHook)
            Dim root As IADRoot = hook.Root
            session = root.Sessions.Item(0)
        Catch ex As Exception
            MessageBox.Show("Could not open Object Viewer: " & ex.Message)
        End Try
    End Sub
    Public Sub OnSelectionChange() Implements IAlibreAddOnCommand.OnSelectionChange
        Try
        Catch ex As Exception
            MessageBox.Show("Error handling selection: " & ex.Message)
        End Try
    End Sub
    Public Sub OnTerminate() Implements IAlibreAddOnCommand.OnTerminate
    End Sub
    Public Function AddTab() As Boolean Implements IAlibreAddOnCommand.AddTab
        Return False
    End Function
    Public Function IsTwoWayToggle() As Boolean Implements IAlibreAddOnCommand.IsTwoWayToggle
        Return False
    End Function
    Public Sub OnShowUI(hWnd As Long) Implements IAlibreAddOnCommand.OnShowUI
    End Sub
    Public Sub OnRender(hDc As Integer, x As Integer, y As Integer, w As Integer, h As Integer) Implements IAlibreAddOnCommand.OnRender
    End Sub
    Public Function OnClick(x As Integer, y As Integer, b As ADDONMouseButtons) As Boolean Implements IAlibreAddOnCommand.OnClick
        Return False
    End Function
    Public Function OnDoubleClick(x As Integer, y As Integer) As Boolean Implements IAlibreAddOnCommand.OnDoubleClick
        Return False
    End Function
    Public Function OnMouseDown(x As Integer, y As Integer, b As ADDONMouseButtons) As Boolean Implements IAlibreAddOnCommand.OnMouseDown
        Return False
    End Function
    Public Function OnMouseMove(x As Integer, y As Integer, b As ADDONMouseButtons) As Boolean Implements IAlibreAddOnCommand.OnMouseMove
        Return False
    End Function
    Public Function OnMouseUp(x As Integer, y As Integer, b As ADDONMouseButtons) As Boolean Implements IAlibreAddOnCommand.OnMouseUp
        Return False
    End Function
    Public Function OnEscape() As Boolean Implements IAlibreAddOnCommand.OnEscape
        Return False
    End Function
    Public Function OnKeyDown(keycode As Integer) As Boolean Implements IAlibreAddOnCommand.OnKeyDown
        Return False
    End Function
    Public Function OnKeyUp(keycode As Integer) As Boolean Implements IAlibreAddOnCommand.OnKeyUp
        Return False
    End Function
    Public Function OnMouseWheel(delta As Double) As Boolean Implements IAlibreAddOnCommand.OnMouseWheel
        Return False
    End Function
    Public Sub On3DRender() Implements IAlibreAddOnCommand.On3DRender : End Sub
    Public ReadOnly Property TabName As String Implements IAlibreAddOnCommand.TabName
        Get
            Return ""
        End Get
    End Property
    Public ReadOnly Property Extents As Array Implements IAlibreAddOnCommand.Extents
        Get
            Return Nothing
        End Get
    End Property
End Class
