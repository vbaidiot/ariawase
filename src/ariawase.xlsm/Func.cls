VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Func"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_ARITY = 4 - 1

Private xxInit As Boolean

Private xObject As Object

Private xCommand As String
Private xReturnType As Integer
Private xProcName As String
Private xCallType As VbCallType

Private xParams As Variant
Private xArity As Integer

Private Sub Bet(ByRef var As Variant, ByRef val As Variant)
    If IsObject(val) Then Set var = val Else Let var = val
End Sub

Private Function Incr(ByRef i As Integer) As Integer
    Incr = i: i = i + 1
End Function

Public Sub Init( _
    ByVal obj As Object, ByVal cmd As String, Optional ByVal params As Variant _
    )
    
    If xxInit Then Err.Raise 5
    
    Set xObject = obj
    
    If Len(cmd) < 4 Then Err.Raise 5
    xCommand = cmd
    Select Case Left(xCommand, 1)
        Case "-":  xReturnType = 0
        Case "+":  xReturnType = 1
        Case "*":  xReturnType = 2
        Case Else: Err.Raise 5
    End Select
    Select Case Right(xCommand, 2)
        Case "<-": xCallType = IIf(xReturnType = 2, VbSet, VbLet)
        Case "->": xCallType = VbGet
        Case "()": xCallType = VbMethod
        Case Else: Err.Raise 5
    End Select
    xProcName = Mid(xCommand, 2, Len(xCommand) - 3)
    
    If IsMissing(params) Then params = Array()
    If Not IsArray(params) Then Err.Raise 13
    xParams = params
    
    xArity = UBound(xParams)
    If xArity > MAX_ARITY Then Err.Raise 5
    
    xxInit = True
End Sub

Public Sub Apply(Optional ByVal params As Variant, Optional ByRef ret As Variant)
    If IsMissing(params) Then params = Array()
    If Not IsArray(params) Then Err.Raise 13
    
    Dim obj As Object
    Dim i As Integer: i = 0
    If Not xObject Is Nothing Then
        If xArity <> UBound(params) Then Err.Raise 5
        Set obj = xObject
    Else
        If xArity <> UBound(params) - 1 Then Err.Raise 5
        Set obj = params(Incr(i))
    End If
    
    Dim a0 As Variant, a1 As Variant, a2 As Variant, a3 As Variant
    If xArity >= 0 Then If IsMissing(params(i)) Then Bet a0, xParams(Incr(i)) Else Bet a0, params(Incr(i))
    If xArity >= 1 Then If IsMissing(params(i)) Then Bet a1, xParams(Incr(i)) Else Bet a1, params(Incr(i))
    If xArity >= 2 Then If IsMissing(params(i)) Then Bet a2, xParams(Incr(i)) Else Bet a2, params(Incr(i))
    If xArity >= 3 Then If IsMissing(params(i)) Then Bet a3, xParams(Incr(i)) Else Bet a3, params(Incr(i))
    
    ret = Missing
    Select Case xReturnType
    Case 0
        Select Case xArity
        Case -1: Call CallByName(obj, xProcName, xCallType)
        Case 0:  Call CallByName(obj, xProcName, xCallType, a0)
        Case 1:  Call CallByName(obj, xProcName, xCallType, a0, a1)
        Case 2:  Call CallByName(obj, xProcName, xCallType, a0, a1, a2)
        Case 3:  Call CallByName(obj, xProcName, xCallType, a0, a1, a2, a3)
        Case Else: Err.Raise 5
        End Select
    Case 1
        Select Case xArity
        Case -1: Let ret = CallByName(obj, xProcName, xCallType)
        Case 0:  Let ret = CallByName(obj, xProcName, xCallType, a0)
        Case 1:  Let ret = CallByName(obj, xProcName, xCallType, a0, a1)
        Case 2:  Let ret = CallByName(obj, xProcName, xCallType, a0, a1, a2)
        Case 3:  Let ret = CallByName(obj, xProcName, xCallType, a0, a1, a2, a3)
        Case Else: Err.Raise 5
        End Select
    Case 2
        Select Case xArity
        Case -1: Set ret = CallByName(obj, xProcName, xCallType)
        Case 0:  Set ret = CallByName(obj, xProcName, xCallType, a0)
        Case 1:  Set ret = CallByName(obj, xProcName, xCallType, a0, a1)
        Case 2:  Set ret = CallByName(obj, xProcName, xCallType, a0, a1, a2)
        Case 3:  Set ret = CallByName(obj, xProcName, xCallType, a0, a1, a2, a3)
        Case Else: Err.Raise 5
        End Select
    Case Else: Err.Raise 51 'Unreachable
    End Select
End Sub
