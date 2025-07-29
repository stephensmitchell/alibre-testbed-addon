import sys
import clr
clr.AddReference("alibre-api")
clr.AddReference("AlibreScriptAddOn")
clr.AddReference('System.Runtime.InteropServices')
from System.Runtime.InteropServices import Marshal
from com.alibre.automation import *
from AlibreScript.API import *
alibre = Marshal.GetActiveObject("AlibreX.AutomationHook")
root = alibre.Root
myPart = Part(root.TopmostSession)
print(myPart)
print(dir(myPart))
print(myPart.FileName)
print(dir(Windows))
Win = Windows()
Win.InfoDialog('This code is from Alibre Script', myPart.FileName)