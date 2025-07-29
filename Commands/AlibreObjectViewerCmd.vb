Imports System.Windows.Forms
Imports AlibreX
Imports AlibreAddOn
Public Class AlibreAddOnCommand
    Implements IAlibreAddOnCommand
    Private session As IADSession
    Public Sub New()
    End Sub
    Public Sub OnComplete() Implements IAlibreAddOnCommand.OnComplete
        Dim hook As IAutomationHook = CType(Runtime.InteropServices.Marshal.GetActiveObject("AlibreX.AutomationHook"), IAutomationHook)
        Dim root As IADRoot = hook.Root
        session = root.Sessions.Item(0)
    End Sub
    Public Sub OnSelectionChange() Implements IAlibreAddOnCommand.OnSelectionChange
        Try
            If session.SelectedObjects.Count = 1 Then
                Dim proxy As IADTargetProxy = CType(session.SelectedObjects.Item(0), IADTargetProxy)
                Dim target As Object = proxy.Target
            End If
        Catch ex As Exception
            MessageBox.Show("Error on selection: " & ex.Message)
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
    Public Sub OnRender(hDc As Integer, clipRectX As Integer, clipRectY As Integer, clipRectWidth As Integer, clipRectHeight As Integer) Implements IAlibreAddOnCommand.OnRender
    End Sub
    Public Function OnClick(screenX As Integer, screenY As Integer, buttons As ADDONMouseButtons) As Boolean Implements IAlibreAddOnCommand.OnClick
        Return False
    End Function
    Public Function OnDoubleClick(screenX As Integer, screenY As Integer) As Boolean Implements IAlibreAddOnCommand.OnDoubleClick
        Return False
    End Function
    Public Function OnMouseDown(screenX As Integer, screenY As Integer, buttons As ADDONMouseButtons) As Boolean Implements IAlibreAddOnCommand.OnMouseDown
        Return False
    End Function
    Public Function OnMouseMove(screenX As Integer, screenY As Integer, buttons As ADDONMouseButtons) As Boolean Implements IAlibreAddOnCommand.OnMouseMove
        Return False
    End Function
    Public Function OnMouseUp(screenX As Integer, screenY As Integer, buttons As ADDONMouseButtons) As Boolean Implements IAlibreAddOnCommand.OnMouseUp
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
    Public Sub On3DRender() Implements IAlibreAddOnCommand.On3DRender
    End Sub
    Public Property CommandSite As IADAddOnCommandSite Implements IAlibreAddOnCommand.CommandSite
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
