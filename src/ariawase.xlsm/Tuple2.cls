VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tuple2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private xxInit As Boolean
Private xxItem1 As Variant
Private xxItem2 As Variant

Public Property Get Item1() As Variant
    If IsObject(xxItem1) Then Set Item1 = xxItem1 Else Let Item1 = xxItem1
End Property

Public Property Get Item2() As Variant
    If IsObject(xxItem2) Then Set Item2 = xxItem2 Else Let Item2 = xxItem2
End Property

Public Sub Init(ByVal itm1 As Variant, ByVal itm2 As Variant)
    If xxInit Then Err.Raise 5
    
    If IsObject(itm1) Then Set xxItem1 = itm1 Else Let xxItem1 = itm1
    If IsObject(itm2) Then Set xxItem2 = itm2 Else Let xxItem2 = itm2
    xxInit = True
End Sub

Public Function Equals(ByVal tpl As Tuple2) As Boolean
    Equals = Core.Equals(xxItem1, tpl.Item1) And Core.Equals(xxItem2, tpl.Item2)
End Function

Public Function Compare(ByVal tpl As Tuple2) As Integer
    Compare = Core.Compare(xxItem1, tpl.Item1)
    If Compare = 0 Then Compare = Core.Compare(xxItem2, tpl.Item2)
End Function

Public Function ToStr() As String
    ToStr = "Tuple2 { Item1 = " & Core.ToStr(xxItem1) & ", Item2 = " & Core.ToStr(xxItem2) & " }"
End Function
