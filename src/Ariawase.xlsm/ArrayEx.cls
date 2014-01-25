VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private xItems As Variant
Private xLength As Long
Private xIndex As Long

Private Sub Class_Initialize()
    xIndex = -1
    xLength = 32
    ReDim xItems(xLength - 1)
End Sub

Private Sub Extend()
    If xIndex < xLength Then GoTo Escape
    xLength = xLength * 2
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
