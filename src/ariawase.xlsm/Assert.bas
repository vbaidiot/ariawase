Attribute VB_Name = "Assert"
Option Explicit

Private xxStartTime As Single
Private xxEndTime As Single
Private xxSuccSubCount As Long
Private xxFailSubCount As Long

Private xxTestIdx As Long
Private xxFailMsgs As Collection

Public Sub TestStart()
    xxSuccSubCount = 0
    xxFailSubCount = 0
    xxStartTime = Timer
End Sub

Public Sub TestEnd()
    xxEndTime = Timer
    
    Debug.Print "===="
    Debug.Print Formats( _
        "{0} succeeded, {1} failed, took {2:0.00} seconds.", _
        xxSuccSubCount, xxFailSubCount, xxEndTime - xxStartTime)
End Sub

Public Sub TestSub(ByVal obj As Object, ByVal proc As String)
    xxTestIdx = 1
    Set xxFailMsgs = New Collection
    
    CallByName obj, proc, VbMethod
    
    If xxFailMsgs.Count < 1 Then
        Debug.Print "+ " & proc
        IncrPre xxSuccSubCount
    Else
        Debug.Print "- " & proc
        Debug.Print "  " & Join(ClctToArr(xxFailMsgs), vbCrLf & "  ")
        IncrPre xxFailSubCount
    End If
End Sub

Private Sub TestDone(ByVal flg As Boolean, ByVal msg As String)
    If Not flg Then Push xxFailMsgs, Formats("[{0}] {1}", xxTestIdx, msg)
    IncrPre xxTestIdx
End Sub

Public Sub IsNullVal(ByVal x As Variant, Optional ByVal msg As String = "")
    TestDone IsNull(x), msg
End Sub

Public Sub IsNotNullVal(ByVal x As Variant, Optional ByVal msg As String = "")
    TestDone Not IsNull(x), msg
End Sub

Public Sub AreEqVal( _
    ByVal ext As Variant, ByVal act As Variant, Optional ByVal msg As String = "" _
    )
    
    TestDone Eq(ext, act), msg
End Sub

Public Sub AreNotEqVal( _
    ByVal ext As Variant, ByVal act As Variant, Optional ByVal msg As String = "" _
    )
    
    TestDone Not Eq(ext, act), msg
End Sub

Public Sub AreEqualVal( _
    ByVal ext As Variant, ByVal act As Variant, Optional ByVal msg As String = "" _
    )
    
    TestDone Equals(ext, act), msg
End Sub

Public Sub AreNotEqualVal( _
    ByVal ext As Variant, ByVal act As Variant, Optional ByVal msg As String = "" _
    )
    
    TestDone Not Equals(ext, act), msg
End Sub

Public Sub AreEqualArr( _
    ByVal ext As Variant, ByVal act As Variant, Optional ByVal msg As String = "" _
    )
    
    TestDone ArrEquals(ext, act), msg
End Sub

Public Sub AreNotEqualArr( _
    ByVal ext As Variant, ByVal act As Variant, Optional ByVal msg As String = "" _
    )
    
    TestDone Not ArrEquals(ext, act), msg
End Sub

Public Sub IsErrFunc( _
    ByVal errnum As Variant, _
    ByVal fun As Func, ByVal params As Variant, _
    Optional ByVal msg As String = "" _
    )
    
    If Not (IsEmpty(errnum) Or IsNumeric(errnum)) Then Err.Raise 5
    If Not IsArray(params) Then Err.Raise 5
    
    On Error GoTo Catch
    
    Dim buf As Variant, ret As Boolean
    fun.CallByPtr buf, params
    TestDone ret, msg
    GoTo Escape
    
Catch:
    ret = IsEmpty(errnum) Or Err.Number = errnum
    Resume Next
    
Escape:
End Sub

Public Sub IsErrMethod( _
    ByVal errnum As Variant, _
    ByVal obj As Object, ByVal proc As String, ByVal params As Variant, _
    Optional ByVal msg As String = "" _
    )
    
    If Not (IsEmpty(errnum) Or IsNumeric(errnum)) Then Err.Raise 5
    If Not IsArray(params) Then Err.Raise 5
    If LBound(params) <> 0 Then Err.Raise 5
    
    On Error GoTo Catch
    
    Dim ret As Boolean
    Select Case UBound(params)
        Case 0:  CallByName obj, proc, VbMethod, params(0)
        Case 1:  CallByName obj, proc, VbMethod, params(0), params(1)
        Case 2:  CallByName obj, proc, VbMethod, params(0), params(1), params(2)
        Case 3:  CallByName obj, proc, VbMethod, params(0), params(1), params(2), params(3)
        Case 4:  CallByName obj, proc, VbMethod, params(0), params(1), params(2), params(3), params(4)
        Case 5:  CallByName obj, proc, VbMethod, params(0), params(1), params(2), params(3), params(4), params(5)
        Case 6:  CallByName obj, proc, VbMethod, params(0), params(1), params(2), params(3), params(4), params(5), params(6)
        Case 7:  CallByName obj, proc, VbMethod, params(0), params(1), params(2), params(3), params(4), params(5), params(6), params(7)
        Case Else: Err.Raise 5
    End Select
    TestDone ret, msg
    GoTo Escape
    
Catch:
    ret = IsEmpty(errnum) Or Err.Number = errnum
    Resume Next
    
Escape:
End Sub
