Imports alibre
Imports AlibreAddOn
Imports AlibreX
Imports IronPython.Hosting
Imports Microsoft.Scripting.Hosting
Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Reflection
Imports System.Windows
Imports IStream = System.Runtime.InteropServices.ComTypes.IStream
Imports MessageBox = System.Windows.MessageBox
Namespace AlibreAddOnAssembly
 Public Module AlibreAddOn
        Private Property AlibreRoot As IADRoot
        Private Property TheAddOnInterface As AddOnRibbon
        Public Sub AddOnLoad(ByVal hwnd As IntPtr, ByVal pAutomationHook As IAutomationHook, ByVal unused As IntPtr)
            AlibreRoot = CType(pAutomationHook.Root, IADRoot)
            TheAddOnInterface = New AddOnRibbon(AlibreRoot)
        End Sub
        Public Sub AddOnUnload(ByVal hwnd As IntPtr, ByVal forceUnload As Boolean, ByRef cancel As Boolean, ByVal reserved1 As Integer, ByVal reserved2 As Integer)
            TheAddOnInterface = Nothing
            AlibreRoot = Nothing
        End Sub
        Public Sub AddOnInvoke(ByVal pAutomationHook As IntPtr, ByVal sessionName As String, ByVal isLicensed As Boolean, ByVal reserved1 As Integer, ByVal reserved2 As Integer)
        End Sub
        Public Function GetAddOnInterface() As IAlibreAddOn
            Return TheAddOnInterface
        End Function
    End Module
End Namespace
