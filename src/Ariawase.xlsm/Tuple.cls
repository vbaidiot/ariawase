VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tuple"
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

Private xInit As Boolean
Private xItems As Variant
Private xLength As Long

Public Property Get Item1() As Variant
    Dim i As Long: i = 0
    If IsObject(xItems(i)) Then Set Item1 = xItems(i) Else Let Item1 = xItems(i)
End Property

Public Property Get Item2() As Variant
    Dim i As Long: i = 1
    If IsObject(xItems(i)) Then Set Item2 = xItems(i) Else Let Item2 = xItems(i)
End Property

Public Property Get Item3() As Variant
    Dim i As Long: i = 2
    If xLength <= i Then Err.Raise 380
    If IsObject(xItems(i)) Then Set Item3 = xItems(i) Else Let Item3 = xItems(i)
End Property

Public Property Get Item4() As Variant
    Dim i As Long: i = 3
    If xLength <= i Then Err.Raise 380
    If IsObject(xItems(i)) Then Set Item4 = xItems(i) Else Let Item4 = xItems(i)
End Property

Public Sub Init(ParamArray itms() As Variant)
    If xInit Then Err.Raise 5
    xItems = itms
    xLength = UBound(itms) + 1
    
    If xLength < 2 Then Err.Raise 5
    xInit = True
End Sub

Public Function ToArray() As Variant
    ToArray = xItems
End Function

Public Function Equals(ByVal tpl As Tuple) As Boolean
    Equals = Core.ArrEquals(xItems, tpl.ToArray())
End Function

Public Function Compare(ByVal tpl As Tuple) As Integer
    Compare = Core.ArrCompare(xItems, tpl.ToArray())
End Function

Public Function ToStr() As String
    Dim s As String, i As Long
    
    s = "Tuple`" & xLength & " { Item1 = " & Dump(xItems(0))
    For i = 1 To xLength - 1
        s = s & ", Item" & i + 1 & " = " & Dump(xItems(i))
    Next
    s = s & " }"
    
    ToStr = s
End Function
