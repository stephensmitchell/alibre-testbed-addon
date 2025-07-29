Imports AlibreAddOn
Imports AlibreX
Imports LINQPad
Imports XDMessaging
Public Module TestbedLink
    Public Const HeaderId As Integer = 9301
    Public Const FileBtnAId As Integer = 9302
    Public Header As String = "TestBed Link"
    Public FileBtnA As String = "TestBed Link"
    Public _HeaderId As Integer()
    Public FileBtnDescrA As String = "TestBed Link"
    Public broadcast As IXDBroadcaster
    Public client As XDMessagingClient
    Public listener As IXDListener
    Public uniqueInstanceName As String
    Public pos As Integer = 0
    Public Const Key_01 As String = "a;b;c;d"
    Public Const Key_02 As String = "a1;b1;c1;d1"
    Public Const Key_03 As String = "0"
    Sub New()
        MsgBox("{CLICK-TEST}")
        client = New XDMessagingClient()
        broadcast = client.Broadcasters.GetBroadcasterForMode(XDTransportMode.HighPerformanceUI, False)
        listener = client.Listeners.GetListenerForMode(XDTransportMode.HighPerformanceUI)
        AddHandler listener.MessageReceived, AddressOf OnMessageReceived
        OnLoad()
        Dim button = New LINQPad.Controls.Button("asdfasdf").Dump()
        AddHandler button.Click, AddressOf SendMessage
        MsgBox("{CLICK-TEST}")
    End Sub
    Public Sub Build()
        _HeaderId = New Integer() {FileBtnAId}
    End Sub
    Public Function FileMenuCmdA(session As IADSession) As IAlibreAddOnCommand
        MsgBox("{CLICK-TEST}")
        Dim message = "a1;b1;c1;d1"
        Dim dumpmsg As New DumpMessage(message)
        broadcast.SendToChannel("Channel1", message)
        MsgBox("{CLICK-TEST}")
        Return Nothing
    End Function
    Private Sub SendMessage()
        Dim message = "a1;b1;c1;d1"
        Dim dumpmsg As New DumpMessage(message)
        broadcast.SendToChannel("Channel1", message)
    End Sub
    Sub OnLoad()
        Console.WriteLine(Util.HostProcessID)
        Debug.WriteLine(Util.HostProcessID)
        uniqueInstanceName = $"{Environment.MachineName}-{Util.HostProcessID}"
        Dim message = $"{uniqueInstanceName} has joined"
        listener.RegisterChannel("Status")
        Dim message1 = $"{uniqueInstanceName} is registering Status."
        broadcast.SendToChannel("Status", message1)
        listener.RegisterChannel("Channel1")
        Dim message2 = $"{uniqueInstanceName} is registering Channel1."
        broadcast.SendToChannel("Channel1", message2)
    End Sub
    Sub OnMessageReceived(ByVal sender As Object, ByVal e As XDMessageEventArgs)
        If e.DataGram.Channel = "Channel1" Then
            Debug.WriteLine(e.DataGram.Message & "::::Channel1")
            MsgBox(e.DataGram.Message & "::::Channel1")
        End If
        If e.DataGram.Channel = "Channel2" Then
            Debug.WriteLine(e.DataGram.Message & "::::Channel2")
            MsgBox(e.DataGram.Message & "::::Channel2")
        End If
        If e.DataGram.Channel = "Status" Then
            Debug.WriteLine(e.DataGram.Message & "::::Status")
            MsgBox(e.DataGram.Message & "::::Status")
        End If
    End Sub
End Module
<Serializable>
Public Class FormattedUserMessage
    Public Sub New(ByVal formatMessage As String, ByVal args As String())
        FormattedTextMessage = String.Format("{0} {1}", formatMessage, args)
    End Sub
    Public ReadOnly Property FormattedTextMessage As String
End Class
<Serializable>
Public Class DumpMessage
    Public Sub New(ByVal formatMessage As String)
        FormattedTextMessage = String.Format("{0} :::::::: {0}", formatMessage)
    End Sub
    Public ReadOnly Property FormattedTextMessage As String
End Class
