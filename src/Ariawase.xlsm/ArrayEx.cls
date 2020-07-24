VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''+----                                                                   --+
'''|                             Ariawase 0.9.0                              |
'''|                Ariawase is free library for VBA cowboys.                |
'''|          The Project Page: https://github.com/vbaidiot/Ariawase         |
'''+--                                                                   ----+
Option Explicit

Private xItems As Variant
Private xIndex As Long

Private Sub Class_Initialize()
    xIndex = -1
    ReDim xItems(31)
End Sub

Public Function AddVal(ByVal val As Variant)
    ArrResizeLet xItems, IncrPre(xIndex), val
End Function

Public Function AddObj(ByVal obj As Variant)
    ArrResizeSet xItems, IncrPre(xIndex), obj
End Function

Public Function ToArray() As Variant
    If xIndex > -1 Then
        ReDim Preserve xItems(xIndex)
        ToArray = xItems
    Else
        ToArray = Array()
    End If
End Function
