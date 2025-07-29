Imports AlibreAddOn
Imports AlibreX
Imports System.Windows.Forms
Module ViewerCommands
    Public Function LaunchObjectViewer(session As IADSession) As IAlibreAddOnCommand
        Try
            Dim targets As New List(Of Object)()
            For i As Integer = 0 To session.SelectedObjects.Count - 1
                Dim proxy = TryCast(session.SelectedObjects.Item(i), IADTargetProxy)
                If proxy IsNot Nothing AndAlso proxy.Target IsNot Nothing Then
                    targets.Add(proxy.Target)
                End If
            Next
            If targets.Count = 0 Then
                MessageBox.Show("No Alibre objects are selected.", "Object Viewer")
                Return Nothing
            End If
        Catch ex As Exception
            MessageBox.Show("Error launching viewer: " & ex.Message)
        End Try
        Return Nothing
    End Function
End Module
