Attribute VB_Name = "Assert"
Option Explicit

Private Const TestClassSuffix As String = "Test"
Private Const TestProcSuffix As String = "_Test"

Private xxStartTime As Single
Private xxEndTime As Single
Private xxSuccSubCount As Long
Private xxFailSubCount As Long

Private xxAssertIx As Long
Private xxFailMsgs As Collection

Private Property Get VBProject() As Object
    Select Case Application.Name
        Case "Microsoft Word":   Set VBProject = Application.MacroContainer.VBProject
        Case "Microsoft Excel":  Set VBProject = Application.ThisWorkbook.VBProject
        Case "Microsoft Access": Set VBProject = Application.VBE.ActiveVBProject
        Case Else: Err.Raise 17
    End Select
End Property

Private Sub TestStart()
    xxSuccSubCount = 0
    xxFailSubCount = 0
    xxStartTime = Timer
End Sub

Private Sub TestEnd()
    xxEndTime = Timer
    
    Debug.Print "===="
    Debug.Print Formats( _
        "{0} succeeded, {1} failed, took {2:0.00} seconds.", _
        xxSuccSubCount, xxFailSubCount, xxEndTime - xxStartTime)
End Sub

Private Sub RunTestSub(ByVal obj As Object, ByVal proc As String)
    xxAssertIx = 1
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

Public Sub RunTestClass(ByVal clsObj As Object, ByVal clsName As String)
    Dim vbcompo As Object: Set vbcompo = VBProject.VBComponents(clsName)
    If vbcompo.Type <> 2 Then Err.Raise 5
    If Right(clsName, Len(TestClassSuffix)) <> TestClassSuffix Then Err.Raise 5
    
    Dim cdmdl As Object:     Set cdmdl = vbcompo.CodeModule
    Dim procs As Collection: Set procs = New Collection
    Dim proc As Variant:     proc = ""
    
    Dim i As Long
    For i = cdmdl.CountOfDeclarationLines To cdmdl.CountOfLines
        If proc <> cdmdl.ProcOfLine(i, 0) Then
            proc = cdmdl.ProcOfLine(i, 0)
            If Right(proc, Len(TestProcSuffix)) = TestProcSuffix Then procs.Add proc
        End If
    Next
    
    TestStart
    For Each proc In procs: RunTestSub clsObj, proc: Next
    TestEnd
End Sub

Private Sub AssertDone(ByVal flg As Boolean, ByVal msg As String)
    If Not flg Then Push xxFailMsgs, Formats("[{0}] {1}", xxAssertIx, msg)
    IncrPre xxAssertIx
End Sub

Public Sub IsNullVal(ByVal x As Variant, Optional ByVal msg As String = "")
    AssertDone IsNull(x), msg
End Sub

Public Sub IsNotNullVal(ByVal x As Variant, Optional ByVal msg As String = "")
    AssertDone Not IsNull(x), msg
End Sub

Public Sub IsInstanceOfTypeName( _
    ByVal expType As String, ByVal x As Variant, Optional ByVal msg As String = "" _
    )
    
    AssertDone TypeName(x) = expType, msg
End Sub

Public Sub IsNotInstanceOfTypeName( _
    ByVal expType As String, ByVal x As Variant, Optional ByVal msg As String = "" _
    )
    
    AssertDone Not TypeName(x) = expType, msg
End Sub

Public Sub AreEqVal( _
    ByVal exp As Variant, ByVal act As Variant, Optional ByVal msg As String = "" _
    )
    
    AssertDone Eq(exp, act), msg
End Sub

Public Sub AreNotEqVal( _
    ByVal exp As Variant, ByVal act As Variant, Optional ByVal msg As String = "" _
    )
    
    AssertDone Not Eq(exp, act), msg
End Sub

Public Sub AreEqualVal( _
    ByVal exp As Variant, ByVal act As Variant, Optional ByVal msg As String = "" _
    )
    
    AssertDone Equals(exp, act), msg
End Sub

Public Sub AreNotEqualVal( _
    ByVal exp As Variant, ByVal act As Variant, Optional ByVal msg As String = "" _
    )
    
    AssertDone Not Equals(exp, act), msg
End Sub

Public Sub AreEqualArr( _
    ByVal exp As Variant, ByVal act As Variant, Optional ByVal msg As String = "" _
    )
    
    AssertDone ArrEquals(exp, act), msg
End Sub

Public Sub AreNotEqualArr( _
    ByVal exp As Variant, ByVal act As Variant, Optional ByVal msg As String = "" _
    )
    
    AssertDone Not ArrEquals(exp, act), msg
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
    AssertDone ret, msg
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
    
    Dim i As Long, ret As Boolean
    Dim ubParam As Long: ubParam = UBound(params)
    Dim ps() As Variant: ReDim ps(ubParam)
    For i = 0 To ubParam: ps(i) = params(i): Next
    
    rtcCallByName obj, StrPtr(proc), VbMethod, ps
    
    AssertDone ret, msg
    GoTo Escape
    
Catch:
    ret = IsEmpty(errnum) Or Err.Number = errnum
    Resume Next
    
Escape:
End Sub
