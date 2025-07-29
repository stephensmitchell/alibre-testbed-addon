Imports AlibreAddOn
Imports AlibreX
Imports System.Threading
Module ToolCommands
    Public Function DummyFunction(session As IADSession) As IAlibreAddOnCommand
        MsgBox("{CLICK-TEST}")
        Return Nothing
    End Function
    Public Async Function CloseAll(session As IADSession) As Task(Of IAlibreAddOnCommand)
        Await Task.Run(Sub() SaveOperation(session))
        Return Nothing
    End Function
    Private Sub SaveOperation(session As IADSession)
        Thread.Sleep(TimeSpan.FromSeconds(5))
        session.Close(True)
        session.Root.TerminateAll()
    End Sub
End Module
