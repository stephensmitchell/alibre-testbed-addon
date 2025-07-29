Imports System.IO
Imports System.Reflection
Imports AlibreAddOn
Imports AlibreX
Module ScriptingCommands
    Public Function RunAddonTestFileCmd(session As IADSession) As IAlibreAddOnCommand
        Dim filePath As String = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "AlibreScripts\sample_alibre_script_000.py")
        Dim eng = IronPython.Hosting.Python.CreateEngine()
        Dim scope = eng.CreateScope()
        Dim scriptFile = File.ReadAllText(filePath)
        eng.Execute(scriptFile, scope)
        Return Nothing
    End Function
    Public Function RunScriptFileCmd() As IAlibreAddOnCommand
        Dim userInput As String
        userInput = InputBox("Enter the path to a Python File:", "Run A Alibre Script Python File", "")
        Dim eng = IronPython.Hosting.Python.CreateEngine()
        Dim scope = eng.CreateScope()
        Dim scriptFile = File.ReadAllText(userInput)
        eng.Execute(scriptFile, scope)
        Return Nothing
    End Function
End Module
