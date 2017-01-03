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
Private xLength As Long
Private xIndex As Long

Private Sub Class_Initialize()
    xIndex = -1
    xLength = -1 + 32
    ReDim xItems(xLength - 1)
End Sub

Private Sub Extend()
    If xIndex < xLength Then GoTo Escape
    xLength = xLength + 1 'possible overflow (Err.Raise 6)
    xLength = -1 + xLength + xLength
    ReDim Preserve xItems(xLength - 1)
    
Escape:
End Sub

Public Function AddVal(ByVal val As Variant)
    xIndex = xIndex + 1
    Extend
    Let xItems(xIndex) = val
End Function

Public Function AddObj(ByVal obj As Variant)
    xIndex = xIndex + 1
    Extend
    Set xItems(xIndex) = obj
End Function

Public Function ToArray() As Variant
    Dim arr As Variant: arr = xItems
    If xIndex > -1 Then
        ReDim Preserve arr(xIndex)
    Else
        arr = Array()
    End If
    ToArray = arr
End Function
