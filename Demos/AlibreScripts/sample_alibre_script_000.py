import sys
import clr
sys.path.append("C:\Program Files\Alibre Design 27.0.1.27039\Program")
sys.path.append("C:\Program Files\Alibre Design 27.0.1.27039\Program\Addons\AlibreScript")
clr.AddReference("AlibreX")
clr.AddReference("alibre-api")
clr.AddReference("AlibreScriptAddOn")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
from System.Windows.Forms import MessageBox, MessageBoxButtons
MessageBox.Show(session.FilePath, "sample_alibre_script_002", MessageBoxButtons.OK)
import AlibreX
import AlibreScript
clr.AddReference('System.Runtime.InteropServices')
from System.Runtime.InteropServices import Marshal
#from com.alibre.automation import *
from AlibreScript.API import *
alibre = Marshal.GetActiveObject("AlibreX.AutomationHook")
root = alibre.Root
session = root.Sessions.Item(0)
objADPartSession = session
print(session.FilePath)
print(objADPartSession.Bodies.Count)
b = objADPartSession.Bodies
verts = b.Item(0).Vertices
print(verts.Count)
def printpoint(x, y, z):
    print("{0} , {1} , {2}".format(x, y, z))
for i in range(verts.Count):
    vert = verts.Item(i)
    point = vert.Point
    printpoint(point.X, point.Y, point.Z)
MessageBox.Show(session.FilePath, "sample_alibre_script_002", MessageBoxButtons.OK)
