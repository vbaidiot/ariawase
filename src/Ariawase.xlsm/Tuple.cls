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
'''|                             Ariawase 0.6.0 Beta                         |
'''|                Ariawase is free library for VBA cowboys.                |
'''|           The Project Page: https://github.com/igeta/Ariawase           |
'''+--                                                                   ----+
Option Explicit

Private xxInit As Boolean
Private xxItems As Variant
Private xxLength As Long

Public Property Get Item1() As Variant
    Dim i As Long: i = 0
    If IsObject(xxItems(i)) Then Set Item1 = xxItems(i) Else Let Item1 = xxItems(i)
End Property

Public Property Get Item2() As Variant
    Dim i As Long: i = 1
    If IsObject(xxItems(i)) Then Set Item2 = xxItems(i) Else Let Item2 = xxItems(i)
End Property

Public Property Get Item3() As Variant
    Dim i As Long: i = 2
    If xxLength <= i Then Err.Raise 380
    If IsObject(xxItems(i)) Then Set Item3 = xxItems(i) Else Let Item3 = xxItems(i)
End Property

Public Property Get Item4() As Variant
    Dim i As Long: i = 3
    If xxLength <= i Then Err.Raise 380
    If IsObject(xxItems(i)) Then Set Item4 = xxItems(i) Else Let Item4 = xxItems(i)
End Property

Public Sub Init(ParamArray itms() As Variant)
    If xxInit Then Err.Raise 5
    xxItems = itms
    xxLength = UBound(itms) + 1
    
    If xxLength < 2 Then Err.Raise 5
    xxInit = True
End Sub

Public Function GetItems() As Variant
    GetItems = xxItems
End Function

Public Function Equals(ByVal tpl As Tuple) As Boolean
    Equals = Core.ArrEquals(xxItems, tpl.GetItems())
End Function

Public Function Compare(ByVal tpl As Tuple) As Integer
    Compare = Core.ArrCompare(xxItems, tpl.GetItems())
End Function

Public Function ToStr() As String
    Dim s As String, i As Long
    
    s = "Tuple`" & xxLength & " { Item1 = " & Core.ToLiteral(xxItems(0))
    For i = 1 To xxLength - 1
        s = s & ", Item" & i + 1 & " = " & Core.ToLiteral(xxItems(i))
    Next
    s = s & " }"
    
    ToStr = s
End Function
