Imports AlibreAddOn
Imports AlibreX
Imports Microsoft.Scripting.Hosting.Shell
Imports System.Collections.Generic
Imports System.IO
Imports System.Reflection
Imports System.Windows.Forms
Namespace alibretestbedaddonForAlibreDesign
    Public Class MenuManager
        Private ReadOnly _rootMenuItem As MenuItem
        Private ReadOnly _menuItems As Dictionary(Of Integer, MenuItem)
        Private ReadOnly _session As IADSession
        Public Sub New(ByVal session As IADSession)
            _rootMenuItem = New MenuItem(401, "Testbed For Alibre Design", "Testbed For Alibre Design")
            _menuItems = New Dictionary(Of Integer, MenuItem)()
            BuildMenus()
        End Sub
        Private Sub BuildMenus()
            Dim G1 As New MenuItem(9000, "{1a}", "{1b}")
            G1.AddSubItem(New MenuItem(9001, "{2a}", "{2b}", "{2c}", AddressOf RunScriptFileCmd))
            G1.AddSubItem(New MenuItem(9002, "{3a}", "{3b}", "{3c}", AddressOf RunAddonTestFileCmd))
            G1.AddSubItem(New MenuItem(9003, "{4a}", "{4b}", "{4c}", AddressOf RunScriptFileCmd))
            G1.AddSubItem(New MenuItem(9004, "{5a}", "{5b}", "{5c}", AddressOf RunAddonTestFileCmd))
            G1.AddSubItem(New MenuItem(9005, "{6a}", "{6b}", "{6c}", AddressOf RunScriptFileCmd))
            G1.AddSubItem(New MenuItem(9006, "{7a}", "{7b}", "{7c}", AddressOf RunAddonTestFileCmd))
            Dim G2 As New MenuItem(9007, "{8a}", "{8b}")
            G2.AddSubItem(New MenuItem(9008, "{9a}", "{9b}", "{9c}", AddressOf RunScriptFileCmd))
            G2.AddSubItem(New MenuItem(9009, "{10a}", "{10b}", "{10c}", AddressOf RunAddonTestFileCmd))
            Dim G3 As New MenuItem(9010, "{11a}", "{11b}")
            G3.AddSubItem(New MenuItem(9014, "{12a}", "{12b}", "{12c}", AddressOf CloseAll))
            Dim G4 As New MenuItem(9015, "{13a}", "{13b}")
            G4.AddSubItem(New MenuItem(9016, "{14a}", "{14b}", "{14c}", AddressOf DummyFunction))
            G4.AddSubItem(New MenuItem(9016, "{15a}", "{15b}", "{15c}", AddressOf DummyFunction))
            Dim G5 As New MenuItem(9018, "{16a}", "{16b}")
            G5.AddSubItem(New MenuItem(9019, "{17a}", "{17b}", "{17c}", AddressOf DummyFunction))
            G5.AddSubItem(New MenuItem(9019, "{18a}", "{18b}", "{18c}", AddressOf DummyFunction))
            Dim G6 As New MenuItem(9021, "{19a}", "{19b}")
            G6.AddSubItem(New MenuItem(9022, "{20a}", "{20b}", "{20c}", AddressOf G3.DummyFunction))
            G6.AddSubItem(New MenuItem(9023, "{21a}", "{21b}", "{21c}", AddressOf G3.DummyFunction))
            G6.AddSubItem(New MenuItem(9024, "{22a}", "{22b}", "{22c}", AddressOf G3.DummyFunction))
            Dim G7 As New MenuItem(9025, "{23a}", "{23b}")
            G7.AddSubItem(New MenuItem(9026, "{24a}", "{24b}", "{24c}", AddressOf G3.DummyFunction))
            G7.AddSubItem(New MenuItem(9027, "{25a}", "{25b}", "{25c}", AddressOf G3.DummyFunction))
            G7.AddSubItem(New MenuItem(9028, "{26a}", "{26b}", "{26c}", AddressOf G3.DummyFunction))
            Dim G8 As New MenuItem(9029, "{27a}", "{27b}")
            G8.AddSubItem(New MenuItem(9030, "{28a}", "{28b}", "{28c}", AddressOf URDFSDFInfoCmd))
            G8.AddSubItem(New MenuItem(90301, "{29a}", "{29b}", "{29c}", AddressOf LaunchObjectViewer))
            Dim G9 As New MenuItem(9031, "{30a}", "{30b}")
            G9.AddSubItem(New MenuItem(9032, "{31a}", "{31b}", "{31c}", AddressOf LoadLP))
            G9.AddSubItem(New MenuItem(9033, "{32a}", "{32b}", "{32c}", AddressOf OpenLocalLP))
            G9.AddSubItem(New MenuItem(9034, "{33a}", "{33b}", "{33c}", AddressOf OpenLocalVSCode))
            G9.AddSubItem(New MenuItem(9035, "{34a}", "{34b}", "{34c}", AddressOf Configuration))
            G9.AddSubItem(New MenuItem(9036, "{35a}", "{35b}", "{35c}", AddressOf TestCmd))
            Dim G10 As New MenuItem(9037, "{36a}", "{36b}")
            G10.AddSubItem(New MenuItem(9038, "{37a}", "{37b}", "{37c}", AddressOf WindowCmd))
            G10.AddSubItem(New MenuItem(9039, "{38a}", "{38b}", "{38c}", AddressOf WindowCmd))
            G10.AddSubItem(New MenuItem(9040, "{39a}", "{39b}", "{39c}", AddressOf WindowCmd))
            Dim G11 As New MenuItem(9041, "{40a}", "{40b}")
            G11.AddSubItem(New MenuItem(9042, "{41a}", "{41b}", "{41c}", AddressOf ParamCmd))
            Dim G12 As New MenuItem(9046, "{42a}", "{42b}")
            G12.AddSubItem(New MenuItem(9047, "{43a}", "{43b}", "{43c}", AddressOf DummyFunction))
            G12.AddSubItem(New MenuItem(9048, "{44a}", "{44b}", "{44c}", AddressOf DummyFunction))
            G12.AddSubItem(New MenuItem(9049, "{45a}", "{45b}", "{45c}", AddressOf DummyFunction))
            Dim G13 As New MenuItem(9050, "{46a}", "{46b}")
            G13.AddSubItem(New MenuItem(9051, "{47a}", "{47b}", "{47c}", AddressOf RegenFeature))
            G13.AddSubItem(New MenuItem(9052, "{48a}", "{48b}", "{48c}", AddressOf RegenDeepFeature))
            G13.AddSubItem(New MenuItem(9053, "{49a}", "{49b}", "{49c}", AddressOf RegenAllFeature))
            _rootMenuItem.AddSubItem(G1)
            _rootMenuItem.AddSubItem(G2)
            _rootMenuItem.AddSubItem(G3)
            _rootMenuItem.AddSubItem(G4)
            _rootMenuItem.AddSubItem(G5)
            _rootMenuItem.AddSubItem(G6)
            _rootMenuItem.AddSubItem(G7)
            _rootMenuItem.AddSubItem(G8)
            _rootMenuItem.AddSubItem(G9)
            _rootMenuItem.AddSubItem(G10)
            _rootMenuItem.AddSubItem(G11)
            _rootMenuItem.AddSubItem(G13)
            _rootMenuItem.AddSubItem(G12)
            RegisterMenuItem(_rootMenuItem)
        End Sub
        Private Sub RegisterMenuItem(menuItem As MenuItem)
            _menuItems(menuItem.Id) = menuItem
            For Each subItem In menuItem.SubItems
                RegisterMenuItem(subItem)
            Next
        End Sub
        Public Function GetMenuItemById(id As Integer) As MenuItem
            Return If(_menuItems.ContainsKey(id), _menuItems(id), Nothing)
        End Function
        Public Function GetRootMenuItem() As MenuItem
            Return _rootMenuItem
        End Function
    End Class
End Namespace