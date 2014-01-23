Attribute VB_Name = "Assert"
'''+----                                                                   --+
'''|                             Ariawase 0.6.0 Beta                         |
'''|                Ariawase is free library for VBA cowboys.                |
'''|           The Project Page: https://github.com/igeta/Ariawase           |
'''+--                                                                   ----+
Option Explicit
Option Private Module

Private Const TestClassSuffix As String = "Test"
Private Const TestProcSuffix As String = "_Test"

Private Const AssertModule As String = "Assert"
Private Const GeneratedProc As String = "TestRunner"
Private Const CommentLineInGeneratedProc As Long = 1

Private Const ResultLineLen As Long = 76

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

Private Function ProcNames(ByVal vbcompo As Object) As Collection
    Dim cdmdl As Object:     Set cdmdl = vbcompo.CodeModule
    Dim procs As Collection: Set procs = New Collection
    Dim proc As Variant:     proc = ""
    
    Dim i As Long
    For i = 1 + cdmdl.CountOfDeclarationLines To cdmdl.CountOfLines
        If proc <> cdmdl.ProcOfLine(i, 0) Then
            proc = cdmdl.ProcOfLine(i, 0)
            procs.Add proc
        End If
    Next
    
    Set ProcNames = procs
End Function

Private Sub WriteResult(ByVal res As String)
    Debug.Print res
End Sub

Private Sub TestStart(ByVal clsName As String)
    WriteResult String(ResultLineLen, "-")
    WriteResult clsName
    WriteResult String(ResultLineLen, "-")
    
    xxSuccSubCount = 0
    xxFailSubCount = 0
    xxStartTime = Timer
End Sub

Private Sub TestEnd()
    xxEndTime = Timer
    
    WriteResult String(ResultLineLen, "=")
    WriteResult _
          xxSuccSubCount & " succeeded, " & xxFailSubCount & " failed," _
        & " took " & Format(xxEndTime - xxStartTime, "0.00") & " seconds."
End Sub

Private Function CheckTestProcName(ByVal proc As String) As Boolean
    CheckTestProcName = Right(proc, Len(TestProcSuffix)) = TestProcSuffix
End Function

Private Function CheckTestClassName(ByVal clsName As String) As Boolean
    CheckTestClassName = Right(clsName, Len(TestClassSuffix)) = TestClassSuffix
End Function

Private Sub RunTestSub(ByVal obj As Object, ByVal proc As String)
    xxAssertIx = 1
    Set xxFailMsgs = New Collection
    
    CallByName obj, proc, VbMethod
    
    If xxFailMsgs.Count < 1 Then
        WriteResult "+ " & proc
        IncrPre xxSuccSubCount
    Else
        WriteResult "- " & proc
        WriteResult "  " & Join(ClctToArr(xxFailMsgs), vbCrLf & "  ")
        IncrPre xxFailSubCount
    End If
End Sub

Public Sub RunTestOf(ByVal clsObj As Object)
    Dim clsName As String: clsName = TypeName(clsObj)
    If Not CheckTestClassName(clsName) Then Err.Raise 5
    
    Dim proc As Variant, procs As Collection
    Set procs = ProcNames(VBProject.VBComponents(clsName))
    
    TestStart clsName
    For Each proc In procs: RunTestSub clsObj, proc: Next
    TestEnd
End Sub

Public Sub RunTest()
    Call TestRunner
End Sub

Private Sub TestRunner()
    ''' NOTE: This is auto-generated code - don't modify contents of this procedure with the code editor.
End Sub

Public Sub TestRunnerClear()
    Dim asrt As Object: Set asrt = VBProject.VBComponents(AssertModule).CodeModule
    Dim st0 As Long: st0 = asrt.ProcStartLine(GeneratedProc, 0)
    Dim st1 As Long: st1 = asrt.ProcBodyLine(GeneratedProc, 0)
    Dim cnt As Long: cnt = asrt.ProcCountLines(GeneratedProc, 0)
    
    asrt.DeleteLines _
        st1 + (1 + CommentLineInGeneratedProc), _
        cnt - ((st1 - st0) + 2 + CommentLineInGeneratedProc)
End Sub

Public Sub TestRunnerGenerate()
    Dim asrt As Object: Set asrt = VBProject.VBComponents(AssertModule).CodeModule
    Dim st1 As Long: st1 = asrt.ProcBodyLine(GeneratedProc, 0)
    Dim pos As Long: pos = st1 + (1 + CommentLineInGeneratedProc)
    
    Dim vbcompo As Object, ln As String
    For Each vbcompo In VBProject.VBComponents
        If vbcompo.Type = 2 And CheckTestClassName(vbcompo.Name) Then
            ln = "Assert.RunTestOf New " & vbcompo.Name
            asrt.InsertLines pos, vbTab & ln
            IncrPre pos
        End If
    Next
End Sub

Private Sub AssertDone( _
    ByVal isa As Boolean, ByVal cond As Boolean, ByVal msg As String, ByVal exp As Variant, ByVal act As Variant _
    )
    
    If isa <> cond Then
        Push xxFailMsgs, "[" & xxAssertIx & "] " & msg & ":"
        Push xxFailMsgs, "  Expected: " & IIf(isa, "", "Not ") & "<" & ToLiteral(exp) & ">"
        Push xxFailMsgs, "  But was:  <" & ToLiteral(act) & ">"
    End If
    IncrPre xxAssertIx
End Sub

Public Sub IsNullVal(ByVal x As Variant, Optional ByVal msg As String = "")
    AssertDone True, IsNull(x), msg, Null, x
End Sub

Public Sub IsNotNullVal(ByVal x As Variant, Optional ByVal msg As String = "")
    AssertDone False, IsNull(x), msg, Null, x
End Sub

Public Sub IsInstanceOfTypeName( _
    ByVal expType As String, ByVal x As Variant, Optional ByVal msg As String = "" _
    )
    
    Dim t As String: t = TypeName(x)
    AssertDone True, expType = t, msg, expType, t
End Sub

Public Sub IsNotInstanceOfTypeName( _
    ByVal expType As String, ByVal x As Variant, Optional ByVal msg As String = "" _
    )
    
    Dim t As String: t = TypeName(x)
    AssertDone False, expType = t, msg, expType, t
End Sub

Public Sub AreEqVal( _
    ByVal exp As Variant, ByVal act As Variant, Optional ByVal msg As String = "" _
    )
    
    AssertDone True, Eq(exp, act), msg, exp, act
End Sub

Public Sub AreNotEqVal( _
    ByVal exp As Variant, ByVal act As Variant, Optional ByVal msg As String = "" _
    )
    
    AssertDone False, Eq(exp, act), msg, exp, act
End Sub

Public Sub AreEqualVal( _
    ByVal exp As Variant, ByVal act As Variant, Optional ByVal msg As String = "" _
    )
    
    AssertDone True, Equals(exp, act), msg, exp, act
End Sub

Public Sub AreNotEqualVal( _
    ByVal exp As Variant, ByVal act As Variant, Optional ByVal msg As String = "" _
    )
    
    AssertDone False, Equals(exp, act), msg, exp, act
End Sub

Public Sub AreEqualArr( _
    ByVal exp As Variant, ByVal act As Variant, Optional ByVal msg As String = "" _
    )
    
    AssertDone True, ArrEquals(exp, act), msg, exp, act
End Sub

Public Sub AreNotEqualArr( _
    ByVal exp As Variant, ByVal act As Variant, Optional ByVal msg As String = "" _
    )
    
    AssertDone False, ArrEquals(exp, act), msg, exp, act
End Sub

Public Sub IsErrFunc( _
    ByVal errnum As Variant, _
    ByVal fun As Func, ByVal params As Variant, _
    Optional ByVal msg As String = "" _
    )
    
    If Not (IsEmpty(errnum) Or IsNumeric(errnum)) Then Err.Raise 5
    If Not IsArray(params) Then Err.Raise 5
    
    On Error GoTo Catch
    
    Dim act As Variant: act = Empty
    
    Dim buf As Variant, ret As Boolean
    fun.CallByPtr buf, params
    AssertDone True, ret, msg, errnum, act
    GoTo Escape
    
Catch:
    act = Err.Number
    ret = IsEmpty(errnum) Or act = errnum
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
    
    Dim act As Variant: act = Empty
    
    Dim i As Long, ret As Boolean
    Dim ubParam As Long: ubParam = UBound(params)
    Dim ps() As Variant: ReDim ps(ubParam)
    For i = 0 To ubParam
        If IsObject(params(i)) Then
            Set ps(i) = params(i)
        Else
            Let ps(i) = params(i)
        End If
    Next
    rtcCallByName obj, StrPtr(proc), VbMethod, ps
    
    AssertDone True, ret, msg, errnum, act
    GoTo Escape
    
Catch:
    act = Err.Number
    ret = IsEmpty(errnum) Or act = errnum
    Resume Next
    
Escape:
End Sub
