Imports AlibreAddOn
Imports AlibreX
Imports LINQPad
Module LINQPadCommands
    Public Function LoadLP(session As IADSession) As IAlibreAddOnCommand
        MsgBox("{CLICK-TEST}")
        Return Nothing
    End Function
    Public Function OpenLocalLP(session As IADSession) As IAlibreAddOnCommand
        MsgBox("{CLICK-TEST}")
        Return Nothing
    End Function
    Public Function OpenLocalVSCode(session As IADSession) As IAlibreAddOnCommand
        MsgBox("{CLICK-TEST}")
        Return Nothing
    End Function
    Public Function Configuration(session As IADSession) As IAlibreAddOnCommand
        MsgBox("{CLICK-TEST}")
        Return Nothing
    End Function
    Public Function TestCmd(session As IADSession) As IAlibreAddOnCommand
        MsgBox("{CLICK-TEST}")
        Util.Compile("{}").Run(QueryResultFormat.Text).AsStringAsync()
        Process.Start("{}")
        Return Nothing
    End Function
End Module
