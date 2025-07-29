Imports AlibreAddOn
Imports AlibreTestbedAddon.alibretestbedaddonForAlibreDesign
Imports AlibreX
Public Class AddOnRibbon
    Implements IAlibreAddOn
    Private ReadOnly _menuManager As MenuManager
    Public _AlibreRoot As IADRoot
    Public _parentWinHandle As IntPtr
    Public Sub New(AlibreRoot As IADRoot, parentWinHandle As IntPtr)
        Try
            _AlibreRoot = AlibreRoot
            _parentWinHandle = parentWinHandle
            _menuManager = New MenuManager(_AlibreRoot.TopmostSession)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Public Sub New(alibreRoot As IADRoot)
        _AlibreRoot = alibreRoot
    End Sub
    Public ReadOnly Property RootMenuItem As Integer Implements IAlibreAddOn.RootMenuItem
        Get
            Return _menuManager.GetRootMenuItem().Id
        End Get
    End Property
    Public Function InvokeCommand(menuId As Integer, sessionIdentifier As String) As IAlibreAddOnCommand Implements IAlibreAddOn.InvokeCommand
        Dim session = _AlibreRoot.Sessions.Item(sessionIdentifier)
        Dim menuItem = _menuManager.GetMenuItemById(menuId)
        Return menuItem?.Command?.Invoke(session)
    End Function
    Public Function HasSubMenus(menuId As Integer) As Boolean Implements IAlibreAddOn.HasSubMenus
        Dim menuItem = _menuManager.GetMenuItemById(menuId)
        Return menuItem IsNot Nothing AndAlso menuItem.SubItems.Count > 0
    End Function
    Public Function SubMenuItems(menuId As Integer) As Array Implements IAlibreAddOn.SubMenuItems
        Dim menuItem = _menuManager.GetMenuItemById(menuId)
        If menuItem IsNot Nothing Then
            Return menuItem.SubItems.Select(Function(subItem) subItem.Id).ToArray()
        End If
        Return Nothing
    End Function
    Public Function MenuItemText(menuId As Integer) As String Implements IAlibreAddOn.MenuItemText
        Dim menuItem = _menuManager.GetMenuItemById(menuId)
        Return menuItem?.Text
    End Function
    Public Function MenuItemState(menuId As Integer, sessionIdentifier As String) As ADDONMenuStates Implements IAlibreAddOn.MenuItemState
        Return ADDONMenuStates.ADDON_MENU_ENABLED
        Process.Start("{48}")
    End Function
    Public Function MenuItemToolTip(menuId As Integer) As String Implements IAlibreAddOn.MenuItemToolTip
        Dim menuItem = _menuManager.GetMenuItemById(menuId)
        Return menuItem?.ToolTip
    End Function
    Public Function MenuIcon(menuID As Integer) As String Implements IAlibreAddOn.MenuIcon
        Dim menuItem = _menuManager.GetMenuItemById(menuID)
        Return menuItem?.Icon
    End Function
    Public Function PopupMenu(menuId As Integer) As Boolean Implements IAlibreAddOn.PopupMenu
        Return False
    End Function
    Public Function HasPersistentDataToSave(sessionIdentifier As String) As Boolean Implements IAlibreAddOn.HasPersistentDataToSave
        Return False
    End Function
    Public Sub SaveData(pCustomData As IStream, sessionIdentifier As String) Implements IAlibreAddOn.SaveData
    End Sub
    Public Sub LoadData(pCustomData As IStream, sessionIdentifier As String) Implements IAlibreAddOn.LoadData
    End Sub
    Public Function UseDedicatedRibbonTab() As Boolean Implements IAlibreAddOn.UseDedicatedRibbonTab
        Return True
    End Function
    Public Sub setIsAddOnLicensed(isLicensed As Boolean)
    End Sub
    Private Sub IAlibreAddOn_setIsAddOnLicensed(isLicensed As Boolean) Implements IAlibreAddOn.setIsAddOnLicensed
    End Sub
End Class
