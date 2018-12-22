Attribute VB_Name = "Core"
'''+----                                                                   --+
'''|                             Ariawase 0.9.0                              |
'''|                Ariawase is free library for VBA cowboys.                |
'''|          The Project Page: https://github.com/vbaidiot/Ariawase         |
'''+--                                                                   ----+
Option Explicit
Option Private Module

''' @seealso WScript.Shell http://msdn.microsoft.com/en-us/library/aew9yb99.aspx (/ja-jp/library/cc364436.aspx)
''' @seealso WbemScripting.SWbemLocator http://msdn.microsoft.com/en-us/library/windows/desktop/aa393719.aspx
''' @seealso VBScript.RegExp http://msdn.microsoft.com/en-us/library/yab2dx62.aspx (/ja-jp/library/cc392403.aspx)

Public Enum HKeyEnum
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    'HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    'HKEY_DYN_DATA = &H80000006
End Enum

Private xxWsh As Object 'Is WScript.Shell
Private xxWmi As Object 'Is WbemScripting.SWbemLocator

Public Property Get Missing() As Variant
    Missing = GetMissing()
End Property
Private Function GetMissing(Optional ByVal mss As Variant) As Variant
    'If Not IsMissing(mss) Then Err.Raise 5
    GetMissing = mss
End Function

''' @param n As Variant
''' @param stepVal As Variant
''' @return As Variant
Public Function IncrPre(ByRef n As Variant, Optional ByVal stepVal As Variant = 1) As Variant
    n = n + stepVal: IncrPre = n
End Function

''' @param n As Variant
''' @param stepVal As Variant
''' @return As Variant
Public Function IncrPst(ByRef n As Variant, Optional ByVal stepVal As Variant = 1) As Variant
    IncrPst = n: n = n + stepVal
End Function

''' @param flgs() As Variant(Of Boolean)
''' @return As Long
Public Function BitFlag(ParamArray flgs() As Variant) As Long
    BitFlag = 0
    Dim ub As Long: ub = UBound(flgs)
    
    Dim i As Long
    For i = 0 To ub
        BitFlag = BitFlag + Abs(flgs(i)) * 2 ^ (ub - i)
    Next
End Function

''' @param num As Variant(Of Numeric Or Date)
''' @return As Boolean
Public Function IsInt(ByVal num As Variant) As Boolean
    If IsDate(num) Then num = CDbl(num)
    If Not IsNumeric(num) Then Err.Raise 13
    
    IsInt = num = Fix(num)
End Function

''' @param num As Variant(Of Numeric)
''' @param digits As Integer
''' @param rndup As Integer
''' @return As Variant(Of Decimal)
Public Function ARound( _
    ByVal num As Variant, Optional ByVal digits As Integer = 0, Optional rndup As Integer = 5 _
    ) As Variant
    
    If Not IsNumeric(num) Then Err.Raise 13
    If Not (1 <= rndup And rndup <= 10) Then Err.Raise 5
    
    num = CDec(num)
    Dim n As Variant: n = CDec(10 ^ Abs(digits))
    Dim z As Variant: z = CDec(Sgn(num) * 0.1 * (10 - rndup))
    If digits >= 0 Then
        ARound = Fix(num * n + z) / n
    Else
        ARound = Fix(num / n + z) * n
    End If
    
Escape:
End Function

''' @param dt As Date
''' @return As Date
Public Function BeginOfMonth(ByVal dt As Date) As Date
    BeginOfMonth = DateAdd("d", -Day(dt) + 1, dt)
End Function

''' @param dt As Date
''' @return As Date
Public Function EndOfMonth(ByVal dt As Date) As Date
    EndOfMonth = DateAdd("d", -1, BeginOfMonth(DateAdd("m", 1, dt)))
End Function

''' @param dt As Date
''' @param fstDayOfWeek As VbDayOfWeek
''' @return As Date
Public Function BeginOfWeek(ByVal dt As Date, Optional fstDayOfWeek As VbDayOfWeek = vbSunday) As Date
    BeginOfWeek = DateAdd("d", 1 - Weekday(dt, fstDayOfWeek), dt)
End Function

''' @param dt As Date
''' @param fstDayOfWeek As VbDayOfWeek
''' @return As Date
Public Function EndOfWeek(ByVal dt As Date, Optional fstDayOfWeek As VbDayOfWeek = vbSunday) As Date
    EndOfWeek = DateAdd("d", 7 - Weekday(dt, fstDayOfWeek), dt)
End Function

'''
''' NOTE: This function for Japanese. Please be customized to your language.
'''
''' @param s As String Is Char
''' @return As Integer
Private Function CharWidth(ByVal s As String) As Integer
   Dim x As Integer: x = Asc(s) / &H100 And &HFF
   CharWidth = 1 + Abs((&H81 <= x And x <= &H9F) Or (&HE0 <= x And x <= &HFC))
End Function

''' @param s As String Is Char
''' @return As Long
Public Function StringWidth(ByVal s As String) As Long
    Dim w As Long: w = 0
    
    Dim i As Long
    For i = 1 To Len(s)
        w = w + CharWidth(Mid$(s, i, 1))
    Next
    StringWidth = w
End Function

''' @param s As String
''' @param byteLen As Long
''' @return As String
Public Function LeftA(ByVal s As String, ByVal byteLen As Long) As String
    Dim ixByte As Long: ixByte = 1
    Dim ixStr As Long:  ixStr = 1
    While (ixByte < 1 + byteLen) And (ixStr <= Len(s))
        ixByte = ixByte + CharWidth(Mid$(s, IncrPst(ixStr), 1))
    Wend
    LeftA = Left$(s, ixStr - (ixByte - byteLen))
End Function

''' @param s As String
''' @param byteLen As Long
''' @return As String
Public Function RightA(ByVal s As String, ByVal byteLen As Long) As String
    Dim idxs As Object: Set idxs = CreateObject("Scripting.Dictionary")
    Dim ixByte As Long: ixByte = 1
    Dim ixStr As Long:  ixStr = 1
    While ixStr <= Len(s)
        idxs.Add ixByte, ixStr
        ixByte = ixByte + CharWidth(Mid$(s, IncrPst(ixStr), 1))
    Wend
    idxs.Add ixByte, ixStr
    
    For byteLen = byteLen To 0 Step -1
        If idxs.Exists(ixByte - byteLen) Then Exit For
    Next
    
    RightA = Right$(s, ixStr - idxs.Item(ixByte - byteLen))
End Function

''' @param s As String
''' @param byteLen As Long
''' @return As Variant(Of Array(Of String))
Public Function SepA(ByVal s As String, ByVal byteLen As Long) As Variant
    Dim ixByte As Long: ixByte = 1
    Dim ixStr  As Long: ixStr = 1
    Dim strLen As Long: strLen = Len(s)
    
    While (ixByte < 1 + byteLen) And (ixStr <= strLen)
        ixByte = ixByte + CharWidth(Mid$(s, IncrPst(ixStr), 1))
    Wend
    
    Dim n As Long: n = ixStr - (ixByte - byteLen)
    SepA = Array(Left$(s, n), Mid$(s, n + 1, strLen))
End Function

''' @param strTemplate As String
''' @param vals() As Variant
''' @return As String
Public Function Formats(ByVal strTemplate As String, ParamArray vals() As Variant) As String
    Dim re As Object: Set re = CreateRegExp("(?:[^\{])?(\{(\d+)(:(.*?[^\}]?))?\})", "g")
    Dim ms As Object: Set ms = re.Execute(strTemplate)
    
    Dim ret As Variant: ret = Array()
    If ms.Count < 1 Then GoTo Ending
    
    ReDim ret(2 * ms.Count)
    Dim ix0 As Long: ix0 = 1
    Dim ix1 As Long: ix1 = 1
    
    Dim i As Long: i = 0
    Dim m As Object, s As String
    For Each m In ms
        ix1 = m.FirstIndex + Abs(Left$(m.Value, 1) <> "{")
        s = Mid$(strTemplate, ix0, ix1 - ix0 + 1)
        Dim mbrc As Variant: mbrc = ReMatch(s, "{+$")
        Dim brcs As String:  If ArrLen(mbrc) > 0 Then brcs = mbrc(0) Else brcs = ""
        
        ret(i + 0) = Replace(Replace(s, "{{", "{"), "}}", "}") 'FIXME: check non-escape brace
        If Len(brcs) Mod 2 = 0 Then
            ret(i + 1) = Format$(vals(m.SubMatches(1)), m.SubMatches(3))
        Else
            ret(i + 1) = m.SubMatches(1)
        End If
        
        i = i + 2
        ix0 = ix1 + Len(m.SubMatches(0)) + 1
    Next
    s = Mid$(strTemplate, ix0)
    ret(i) = Replace(Replace(s, "{{", "{"), "}}", "}") 'FIXME: check non-escape brace
    
Ending:
    Formats = Join(ret, "")
End Function

''' @param obj As Object Is T
''' @param params As Variant()
''' @return As Object Is T
Public Function Init( _
    ByVal obj As Object, _
    Optional ByRef arg1 As Variant, _
    Optional ByRef arg2 As Variant, _
    Optional ByRef arg3 As Variant, _
    Optional ByRef arg4 As Variant, _
    Optional ByRef arg5 As Variant, _
    Optional ByRef arg6 As Variant, _
    Optional ByRef arg7 As Variant, _
    Optional ByRef arg8 As Variant _
    ) As Object
    
    If IsMissing(arg1) Then obj.Init Else _
    If IsMissing(arg2) Then obj.Init arg1 Else _
    If IsMissing(arg3) Then obj.Init arg1, arg2 Else _
    If IsMissing(arg4) Then obj.Init arg1, arg2, arg3 Else _
    If IsMissing(arg5) Then obj.Init arg1, arg2, arg3, arg4 Else _
    If IsMissing(arg6) Then obj.Init arg1, arg2, arg3, arg4, arg5 Else _
    If IsMissing(arg7) Then obj.Init arg1, arg2, arg3, arg4, arg5, arg6 Else _
    If IsMissing(arg8) Then obj.Init arg1, arg2, arg3, arg4, arg5, arg6, arg7 Else _
    obj.Init arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8
    
    Set Init = obj
End Function

''' @param x As Variant
''' @return As String
Public Function ToStr(ByVal x As Variant) As String
    If IsObject(x) Then
        On Error GoTo Err438
        ToStr = x.ToStr()
        On Error GoTo 0
    ElseIf IsArray(x) Then
        ToStr = TypeName(x)
    Else
        ToStr = x
    End If
    
    GoTo Escape
    
Err438:
    Dim e As ErrObject: Set e = Err
    Select Case e.Number
        Case 438: ToStr = TypeName(x): Resume Next
        Case Else: Err.Raise e.Number, e.Source, e.Description, e.HelpFile, e.HelpContext
    End Select
    
Escape:
End Function

''' @param x As Variant
''' @return As String
Public Function Dump(ByVal x As Variant) As String
    Dim ty As String: ty = TypeName(x)
    Select Case ty
    Case "Boolean":     Dump = x
    Case "Integer":     Dump = x & "%"
    Case "Long":        Dump = x & "&"
    #If Win64 Then
    Case "LongLong":    Dump = x & "^"
    #End If
    Case "Single":      Dump = x & "!"
    Case "Double":      Dump = x & "#"
    Case "Currency":    Dump = x & "@"
    Case "Byte":        Dump = "CByte(" & x & ")"
    Case "Decimal":     Dump = "CDec(" & x & ")"
    Case "Date":
        Dim d As String, t As String
        If Abs(x) >= 1 Then d = Month(x) & "/" & Day(x) & "/" & Year(x)
        If Not IsInt(x) Then t = Format(x, "h:nn:ss AM/PM")
        Dump = "#" & Trim(d & " " & t) & "#"
    Case "String"
        If StrPtr(x) = 0 Then
            Dump = "vbNullString"
        Else
            Dump = """" & Replace(x, """", """""") & """"
        End If
    Case "Error"
        If IsMissing(x) Then
            Dump = "Missing"
        Else
            Dump = "CVErr(" & ReMatch(CStr(x), "\d+")(0) & ")"
        End If
    Case "ErrObject"
        Dump = "Err " & x.Number
    Case "Empty", "Null", "Nothing", "Unknown"
        Dump = ty
    Case Else
        If IsObject(x) Then
            Dump = ToStr(x)
        ElseIf IsArray(x) Then
            Dim rnk As Integer: rnk = ArrRank(x)
            If rnk = 1 Then
                Dim lb As Long: lb = LBound(x)
                Dim ub As Long: ub = UBound(x)
                Dim ar As Variant
                If ub - lb < 0 Then
                    ar = Array()
                Else
                    Dim mx As Long: mx = 8 - 1
                    Dim xb As Long: xb = IIf(ub - lb < mx, ub, lb + mx)
                    ReDim ar(lb To xb)
                    Dim i As Long
                    For i = lb To xb: ar(i) = Dump(x(i)): Next
                End If
                Dump = "Array(" & Join(ar, ", ") & IIf(xb < ub, ", ...", "") & ")"
            Else
                Dump = Replace(ty, "()", "(" & String(rnk - 1, ",") & ")")
            End If
        Else
            Err.Raise 51
        End If
    End Select
End Function

''' @param x As Variant(Of T)
''' @param y As Variant(Of T)
''' @return As Variant(Of Boolean Or Null Or Empty)
Public Function Eq(ByVal x As Variant, ByVal y As Variant) As Variant
    Dim xIsObj As Boolean: xIsObj = IsObject(x)
    Dim yIsObj As Boolean: yIsObj = IsObject(y)
    If xIsObj Xor yIsObj Then
        Eq = Empty
    ElseIf xIsObj And yIsObj Then
        Eq = x Is y
    Else
        Eq = x = y
    End If
End Function

''' @param x As Variant(Of T)
''' @param y As Variant(Of T)
''' @param swAllowNull As Boolean
''' @return As Variant(Of Boolean Or Null Or Empty)
Public Function Equals( _
    ByVal x As Variant, ByVal y As Variant, _
    Optional ByVal swAllowNull As Boolean = False _
    ) As Variant
    
    Dim xIsObj As Boolean: xIsObj = IsObject(x)
    Dim yIsObj As Boolean: yIsObj = IsObject(y)
    If xIsObj Or yIsObj Then
        If xIsObj And yIsObj Then Equals = x.Equals(y) Else Equals = Empty
    Else
        Dim xIsNul As Boolean: xIsNul = IsNull(x)
        Dim yIsNul As Boolean: yIsNul = IsNull(y)
        If xIsNul Or yIsNul Then
            If swAllowNull Then Equals = xIsNul And yIsNul Else Equals = Null
        Else
            If TypeName(x) = TypeName(y) Then Equals = x = y Else Equals = Empty
        End If
    End If
End Function

''' @param x As Variant(Of T)
''' @param y As Variant(Of T)
''' @param swAllowNull As Boolean
''' @return As Variant(Of Integer Or Null)
Public Function Compare( _
    ByVal x As Variant, ByVal y As Variant, _
    Optional ByVal swAllowNull As Boolean = False _
    ) As Variant
    
    Dim xIsObj As Boolean: xIsObj = IsObject(x)
    Dim yIsObj As Boolean: yIsObj = IsObject(y)
    If xIsObj Or yIsObj Then
        If xIsObj And yIsObj Then Compare = x.Compare(y) Else Err.Raise 13
    Else
        Dim xIsNul As Boolean: xIsNul = IsNull(x)
        Dim yIsNul As Boolean: yIsNul = IsNull(y)
        If xIsNul Or yIsNul Then
            If swAllowNull Then
                If xIsNul And yIsNul Then Compare = 0 Else _
                If xIsNul Then Compare = -1 Else _
                Compare = 1
            Else
                Compare = Null
            End If
        Else
            If TypeName(x) = TypeName(y) Then
                If x = y Then Compare = 0 Else _
                If x < y Then Compare = -1 Else _
                Compare = 1
            Else
                Err.Raise 13
            End If
        End If
    End If
End Function

Private Sub MinOrMax(ByVal arr As Variant, ByVal comp As Integer, ByRef ret As Variant)
    ret = Empty
    Dim ub As Variant: ub = UBound(arr)
    If ub < 0 Then GoTo Escape
    
    Dim i As Long
    If IsObject(arr(0)) Then
        Set ret = arr(0)
        For i = 1 To ub
            If Compare(arr(i), ret) = comp Then Set ret = arr(i)
        Next
    Else
        Let ret = arr(0)
        For i = 1 To ub
            If Compare(arr(i), ret) = comp Then Let ret = arr(i)
        Next
    End If
    
Escape:
End Sub

''' @param arr() As Variant(Of T)
''' @return As Variant(Of T)
Public Function Min(ParamArray arr() As Variant) As Variant
    MinOrMax arr, -1, Min
End Function

''' @param arr() As Variant(Of T)
''' @return As Variant(Of T)
Public Function Max(ParamArray arr() As Variant) As Variant
    MinOrMax arr, 1, Max
End Function

''' @param arr As Variant(Of Array(Of T))
''' @return As Integer
Public Function ArrRank(ByVal arr As Variant) As Integer
    If Not IsArray(arr) Then Err.Raise 13
    
    Dim x As Long
    Dim i As Integer: i = 0
    On Error Resume Next
    While Err.Number = 0: x = UBound(arr, IncrPre(i)): Wend
    ArrRank = i - 1
End Function

''' @param arr As Variant(Of Array(Of T))
''' @param dimen As Integer
''' @return As Long
Public Function ArrLen(ByVal arr As Variant, Optional ByVal dimen As Integer = 1) As Long
    If Not IsArray(arr) Then Err.Raise 13
    ArrLen = UBound(arr, dimen) - LBound(arr, dimen) + 1
End Function

''' @param arr1 As Variant(Of Array(Of T))
''' @param arr2 As Variant(Of Array(Of T))
''' @param swAllowNull As Boolean
''' @return As Variant(Of Boolean Or Null Or Empty)
Public Function ArrEquals( _
    ByVal arr1 As Variant, ByVal arr2 As Variant, _
    Optional ByVal swAllowNull As Boolean = False _
    ) As Variant
    
    If Not (IsArray(arr1) And IsArray(arr2)) Then Err.Raise 13
    
    Dim lb1 As Long, alen1 As Long: lb1 = LBound(arr1): alen1 = UBound(arr1) - lb1 + 1
    Dim lb2 As Long, alen2 As Long: lb2 = LBound(arr2): alen2 = UBound(arr2) - lb2 + 1
    Dim compLen As Integer: compLen = Compare(alen1, alen2)
    
    Dim ret As Variant: ret = True
    Dim i As Long: i = 0
    Dim alen As Long: alen = IIf(compLen < 1, alen1, alen2)
    While ret And (i < alen)
        ret = Equals(arr1(lb1 + i), arr2(lb2 + i), swAllowNull)
        i = i + 1
    Wend
    If ret Then ret = (compLen = 0)
    ArrEquals = ret
End Function

''' @param arr1 As Variant(Of Array(Of T))
''' @param arr2 As Variant(Of Array(Of T))
''' @return As Variant(Of Integer Or Null)
Public Function ArrCompare( _
    ByVal arr1 As Variant, ByVal arr2 As Variant, _
    Optional ByVal swAllowNull As Boolean = False _
    ) As Variant
    
    If Not (IsArray(arr1) And IsArray(arr2)) Then Err.Raise 13
    
    Dim lb1 As Long, alen1 As Long: lb1 = LBound(arr1): alen1 = UBound(arr1) - lb1 + 1
    Dim lb2 As Long, alen2 As Long: lb2 = LBound(arr2): alen2 = UBound(arr2) - lb2 + 1
    Dim compLen As Integer: compLen = Compare(alen1, alen2)
    
    Dim ret As Variant: ret = 0
    Dim i As Long: i = 0
    Dim alen As Long: alen = IIf(compLen < 1, alen1, alen2)
    While ret = 0 And (i < alen)
        ret = Compare(arr1(lb1 + i), arr2(lb2 + i), swAllowNull)
        i = i + 1
    Wend
    If ret = 0 Then ret = compLen
    ArrCompare = ret
End Function

''' @param arr As Variant(Of Array(Of T))
''' @param val As Variant(Of T)
''' @param ixStart As Variant(Of Long)
''' @param cnt As Variant(Of Long)
''' @return As Long
Public Function ArrIndexOf( _
    ByVal arr As Variant, ByVal val As Variant, _
    Optional ByVal ixStart As Variant, Optional ByVal cnt As Variant _
    ) As Long
    
    If Not IsArray(arr) Then Err.Raise 13
    
    Dim ix0 As Long:  ix0 = LBound(arr)
    Dim alen As Long: alen = ArrLen(arr)
    If IsMissing(ixStart) Then ixStart = ix0
    If IsNumeric(ixStart) Then ixStart = CLng(ixStart) Else Err.Raise 13
    If ixStart < ix0 Then Err.Raise 5
    If IsMissing(cnt) Then cnt = alen
    If IsNumeric(cnt) Then cnt = CLng(cnt) Else Err.Raise 13
    cnt = Min(cnt, alen)
    
    ArrIndexOf = ixStart - 1
    
    Dim i As Long
    For i = ixStart To ixStart + cnt - 1
        If Equals(arr(i), val, True) Then
            ArrIndexOf = i
            GoTo Escape
        End If
    Next
    
Escape:
End Function

''' @param arr As Variant(Of Array(Of T))
Public Sub ArrRev(ByRef arr As Variant)
    If Not IsArray(arr) Then Err.Raise 13
    
    Dim ixL As Long: ixL = LBound(arr)
    Dim ixU As Long: ixU = UBound(arr)
    If ixU - ixL < 1 Then GoTo Escape
    
    Dim x As Variant
    If IsObject(arr(ixL)) Then
        While ixL < ixU
            Set x = arr(ixL): Set arr(ixL) = arr(ixU): Set arr(ixU) = x
            ixL = ixL + 1: ixU = ixU - 1
        Wend
    Else
        While ixL < ixU
            Let x = arr(ixL): Let arr(ixL) = arr(ixU): Let arr(ixU) = x
            ixL = ixL + 1: ixU = ixU - 1
        Wend
    End If
    
Escape:
End Sub

''' @param arr As Variant(Of Array(Of T))
''' @param orderAsc As Boolean
Public Sub ArrSort(ByRef arr As Variant, Optional ByVal orderAsc As Boolean = True)
    If Not IsArray(arr) Then Err.Raise 13
    If ArrLen(arr) <= 1 Then GoTo Escape
    
    Dim ix0 As Long: ix0 = LBound(arr)
    If IsObject(arr(ix0)) Then
        ObjArrMSort arr, ix0, orderAsc
    Else
        ValArrMSort arr, ix0, orderAsc
    End If
    
Escape:
End Sub
Private Sub ObjArrMSort(arr As Variant, lb As Long, orderAsc As Boolean)
    Dim alen As Long: alen = ArrLen(arr)
    If alen <= 1 Then GoTo Escape
    
    '' optimization
    If alen <= 8 Then
        ObjArrISort arr, lb, orderAsc
        GoTo Escape
    End If
    
    Dim i As Long
    Dim l1 As Long: l1 = Fix(alen / 2)
    Dim l2 As Long: l2 = alen - l1
    
    Dim ub1 As Long:   ub1 = lb + l1 - 1
    Dim a1 As Variant: ReDim a1(lb To ub1)
    For i = lb To ub1: Set a1(i) = arr(i): Next
    ObjArrMSort a1, lb, orderAsc
    
    Dim ub2 As Long:   ub2 = lb + l2 - 1
    Dim a2 As Variant: ReDim a2(lb To ub2)
    For i = lb To ub2: Set a2(i) = arr(l1 + i): Next
    ObjArrMSort a2, lb, orderAsc
    
    Dim i1 As Long: i1 = lb
    Dim i2 As Long: i2 = lb
    While i1 <= ub1 Or i2 <= ub2
        If ArrMergeSw(a1, i1, ub1, a2, i2, ub2, orderAsc) Then
            Set arr(i1 + i2 - lb) = a1(IncrPst(i1))
        Else
            Set arr(i1 + i2 - lb) = a2(IncrPst(i2))
        End If
    Wend
    
Escape:
End Sub
Private Sub ValArrMSort(arr As Variant, lb As Long, orderAsc As Boolean)
    Dim alen As Long: alen = ArrLen(arr)
    If alen <= 1 Then GoTo Escape
    
    '' optimization
    If alen <= 8 Then
        ValArrISort arr, lb, orderAsc
        GoTo Escape
    End If
    
    Dim i As Long
    Dim l1 As Long: l1 = Fix(alen / 2)
    Dim l2 As Long: l2 = alen - l1
    
    Dim ub1 As Long:   ub1 = lb + l1 - 1
    Dim a1 As Variant: ReDim a1(lb To ub1)
    For i = lb To ub1: Let a1(i) = arr(i): Next
    ValArrMSort a1, lb, orderAsc
    
    Dim ub2 As Long:   ub2 = lb + l2 - 1
    Dim a2 As Variant: ReDim a2(lb To ub2)
    For i = lb To ub2: Let a2(i) = arr(l1 + i): Next
    ValArrMSort a2, lb, orderAsc
    
    Dim i1 As Long: i1 = lb
    Dim i2 As Long: i2 = lb
    While i1 <= ub1 Or i2 <= ub2
        If ArrMergeSw(a1, i1, ub1, a2, i2, ub2, orderAsc) Then
            Let arr(i1 + i2 - lb) = a1(IncrPst(i1))
        Else
            Let arr(i1 + i2 - lb) = a2(IncrPst(i2))
        End If
    Wend
    
Escape:
End Sub
Private Sub ObjArrISort(arr As Variant, lb As Long, orderAsc As Boolean)
    Dim i As Long, j As Long, x As Variant
    For i = lb + 1 To UBound(arr)
        j = i
        Do While j > lb
            If Compare(arr(j - 1), arr(j), True) * (Abs(orderAsc) * 2 - 1) <= 0 Then Exit Do
            Set x = arr(j): Set arr(j) = arr(j - 1): Set arr(j - 1) = x
            j = j - 1
        Loop
    Next
End Sub
Private Sub ValArrISort(arr As Variant, lb As Long, orderAsc As Boolean)
    Dim i As Long, j As Long, x As Variant
    For i = lb + 1 To UBound(arr)
        j = i
        Do While j > lb
            If Compare(arr(j - 1), arr(j), True) * (Abs(orderAsc) * 2 - 1) <= 0 Then Exit Do
            Let x = arr(j): Let arr(j) = arr(j - 1): Let arr(j - 1) = x
            j = j - 1
        Loop
    Next
End Sub
Private Function ArrMergeSw( _
    arr1 As Variant, i1 As Long, ub1 As Long, _
    arr2 As Variant, i2 As Long, ub2 As Long, _
    orderAsc As Boolean _
    ) As Boolean
    
    If i1 > ub1 Then ArrMergeSw = False Else _
    If i2 > ub2 Then ArrMergeSw = True Else _
    ArrMergeSw = Compare(arr1(i1), arr2(i2), True) * (Abs(orderAsc) * 2 - 1) < 1
End Function

''' @param arr As Variant(Of Array(Of T))
''' @return As Variant(Of Array(Of T))
Public Function ArrUniq(ByVal arr As Variant) As Variant
    If Not IsArray(arr) Then Err.Raise 13
    Dim ret As Variant: ret = Array()
    Dim lbA As Long: lbA = LBound(arr)
    Dim ubA As Long: ubA = UBound(arr)
    If ubA - lbA < 0 Then GoTo Ending
    
    ReDim ret(lbA To ubA)
    
    Dim ixA As Long, ixR As Long: ixR = lbA
    If IsObject(arr(lbA)) Then
        For ixA = lbA To ubA
            If ArrIndexOf(ret, arr(ixA), lbA, ixR - lbA) < lbA Then
                Set ret(IncrPst(ixR)) = arr(ixA)
            End If
        Next
    Else
        For ixA = lbA To ubA
            If ArrIndexOf(ret, arr(ixA), lbA, ixR - lbA) < lbA Then
                Let ret(IncrPst(ixR)) = arr(ixA)
            End If
        Next
    End If
    
    ReDim Preserve ret(lbA To ixR - 1)
    
Ending:
    ArrUniq = ret
End Function

''' @param arr1 As Variant(Of Array(Of T))
''' @param arr2 As Variant(Of Array(Of T))
''' @return As Variant(Of Array(Of T))
Public Function ArrConcat(ByVal arr1 As Variant, ByVal arr2 As Variant) As Variant
    If Not (IsArray(arr1) And IsArray(arr2)) Then Err.Raise 13
    
    Dim lb2 As Long: lb2 = LBound(arr2)
    Dim ub2 As Long: ub2 = UBound(arr2)
    Dim alen2 As Long: alen2 = ub2 - lb2 + 1
    If alen2 < 1 Then GoTo Ending
    
    Dim isObj2 As Boolean: isObj2 = IsObject(arr2(lb2))
    
    Dim lb1 As Long: lb1 = LBound(arr1)
    Dim ub1 As Long: ub1 = UBound(arr1)
    Dim alen1 As Long: alen1 = ub1 - lb1 + 1
    If alen1 > 0 Then If IsObject(arr1(lb1)) <> isObj2 Then Err.Raise 13
    
    Dim ret As Variant: ReDim ret(alen1 + alen2 - 1)
    
    Dim i As Integer
    If isObj2 Then
        For i = 0 To alen1 - 1: Set ret(i) = arr1(lb1 + i): Next
        For i = 0 To alen2 - 1: Set ret(alen1 + i) = arr2(lb2 + i): Next
    Else
        For i = 0 To alen1 - 1: Let ret(i) = arr1(lb1 + i): Next
        For i = 0 To alen2 - 1: Let ret(alen1 + i) = arr2(lb2 + i): Next
    End If
    
Ending:
    ArrConcat = ret
End Function

''' @param arr As Variant(Of Array(Of T))
''' @param ixStart As Variant(Of Long)
''' @param ixEnd As Variant(Of Long)
''' @return As Variant(Of Array(Of T))
Public Function ArrSlice( _
    ByVal arr As Variant, _
    Optional ByVal ixStart As Variant, Optional ByVal ixEnd As Variant _
    ) As Variant
    
    If Not IsArray(arr) Then Err.Raise 13
    
    Dim lbA As Long: lbA = LBound(arr)
    Dim ubA As Long: ubA = UBound(arr)
    If IsMissing(ixStart) Then ixStart = lbA
    If IsNumeric(ixStart) Then ixStart = CLng(ixStart) Else Err.Raise 13
    If IsMissing(ixEnd) Then ixEnd = ubA
    If IsNumeric(ixEnd) Then ixEnd = CLng(ixEnd) Else Err.Raise 13
    
    If Not (lbA <= ixStart And ixEnd <= ubA) Then Err.Raise 5
    
    Dim ret As Variant: ret = Array()
    Dim ubR As Long: ubR = ixEnd - ixStart
    If ubR < 1 Then GoTo Ending
    
    ReDim ret(ubR)
    Dim isObj As Boolean: isObj = IsObject(arr(ixStart))
    
    Dim i As Long
    If isObj Then
        For i = 0 To ubR: Set ret(i) = arr(ixStart + i): Next
    Else
        For i = 0 To ubR: Let ret(i) = arr(ixStart + i): Next
    End If
    
Ending:
    ArrSlice = ret
End Function

''' @param jagArray As Variant(Of Array(Of Array(Of T)))
''' @return As Variant(Of Array(Of T))
Public Function ArrFlatten(ByVal jagArr As Variant) As Variant
    If Not IsArray(jagArr) Then Err.Raise 13
    Dim ret As Variant: ret = Array()
    If ArrLen(jagArr) < 1 Then GoTo Ending
    
    Dim arr As Variant
    For Each arr In jagArr: ret = ArrConcat(ret, arr): Next
    
Ending:
    ArrFlatten = ret
End Function

''' @param clct As Collection(Of T)
''' @param val As Variant(Of T)
Public Sub Push(ByVal clct As Collection, ByVal val As Variant)
    clct.Add val
End Sub

''' @param clct As Collection(Of T)
''' @return As Variant(Of T)
Public Function Pop(ByVal clct As Collection) As Variant
    Dim i As Long: i = clct.Count
    If IsObject(clct.Item(i)) Then Set Pop = clct.Item(i) Else Let Pop = clct.Item(i)
    clct.Remove i
End Function

''' @param clct As Collection(Of T)
''' @param val As Variant(Of T)
Public Sub Shift(ByVal clct As Collection, ByVal val As Variant)
    If clct.Count < 1 Then
        clct.Add val
    Else
        clct.Add val, , 1
    End If
End Sub

''' @param clct As Collection(Of T)
''' @return As Variant(Of T)
Public Function Unshift(ByVal clct As Collection) As Variant
    Dim i As Long: i = 1
    If IsObject(clct.Item(i)) Then Set Unshift = clct.Item(i) Else Let Unshift = clct.Item(i)
    clct.Remove i
End Function

''' @param arr As Variant(Of Array(Of T))
''' @return As Collection(Of T)
Public Function ArrToClct(ByVal arr As Variant) As Collection
    If Not IsArray(arr) Then Err.Raise 13
    Set ArrToClct = New Collection
    Dim v As Variant
    For Each v In arr: ArrToClct.Add v: Next
End Function

''' @param clct As Collection(Of T)
''' @return As Variant(Of Array(Of T))
Public Function ClctToArr(ByVal clct As Collection) As Variant
    Dim arr As Variant: arr = Array()
    Dim clen As Long: clen = clct.Count
    If clen < 1 Then GoTo Ending
    
    ReDim arr(clen - 1)
    Dim i As Long: i = 0
    Dim v As Variant
    If IsObject(clct.Item(1)) Then
        For Each v In clct: Set arr(IncrPst(i)) = v: Next
    Else
        For Each v In clct: Let arr(IncrPst(i)) = v: Next
    End If
    
Ending:
    ClctToArr = arr
End Function

''' @param jagArr As Variant(Of Array(Of Array(Of T))
''' @return As Variant(Of Array(Of T, T))
Public Function JagArrToArr2D(ByVal jagArr As Variant) As Variant
    Dim arr2D As Variant: arr2D = Array()
    
    Dim ixOut As Long, ixInn As Long
    Dim lbOut As Long, lbInn As Long, lbInnFst As Long
    Dim ubOut As Long, ubInn As Long, ubInnFst As Long
    
    If Not IsArray(jagArr) Then Err.Raise 13
    lbOut = LBound(jagArr)
    ubOut = UBound(jagArr)
    If ubOut - lbOut < 0 Then GoTo Ending
    
    If Not IsArray(jagArr(lbOut)) Then Err.Raise 13
    
    lbInnFst = LBound(jagArr(lbOut))
    ubInnFst = UBound(jagArr(lbOut))
    If ubInnFst - lbInnFst < 0 Then
        For ixOut = lbOut + 1 To ubOut
            If ArrLen(jagArr(ixOut)) > 0 Then Err.Raise 5
        Next
        GoTo Ending
    End If
    
    ReDim arr2D(lbOut To ubOut, lbInnFst To ubInnFst)
    If IsObject(jagArr(lbOut)(lbInnFst)) Then
        For ixOut = lbOut To ubOut
            lbInn = LBound(jagArr(ixOut))
            ubInn = UBound(jagArr(ixOut))
            If lbInn <> lbInnFst Or ubInn <> ubInnFst Then Err.Raise 5
            For ixInn = lbInn To ubInn: Set arr2D(ixOut, ixInn) = jagArr(ixOut)(ixInn): Next
        Next
    Else
        For ixOut = lbOut To ubOut
            lbInn = LBound(jagArr(ixOut))
            ubInn = UBound(jagArr(ixOut))
            If lbInn <> lbInnFst Or ubInn <> ubInnFst Then Err.Raise 5
            For ixInn = lbInn To ubInn: Let arr2D(ixOut, ixInn) = jagArr(ixOut)(ixInn): Next
        Next
    End If
    
Ending:
    JagArrToArr2D = arr2D
End Function

''' @param arr2D As Variant(Of Array(Of T, T))
''' @return As Variant(Of Array(Of Array(Of T))
Public Function Arr2DToJagArr(ByVal arr2D As Variant) As Variant
    Dim jagArr As Variant: jagArr = Array()
    
    Dim lb1 As Long, ub1 As Long: lb1 = LBound(arr2D, 1): ub1 = UBound(arr2D, 1)
    If ub1 - lb1 < 0 Then GoTo Ending
    ReDim jagArr(lb1 To ub1)
    
    Dim lb2 As Long, ub2 As Long: lb2 = LBound(arr2D, 2): ub2 = UBound(arr2D, 2)
    Dim ix1 As Long, ix2 As Long
    Dim arr As Variant: ReDim arr(lb2 To ub2)
    
    If IsObject(arr2D(lb1, lb2)) Then
        For ix1 = lb1 To ub1
            jagArr(ix1) = arr
            For ix2 = lb2 To ub2: Set jagArr(ix1)(ix2) = arr2D(ix1, ix2): Next
        Next
    Else
        For ix1 = lb1 To ub1
            jagArr(ix1) = arr
            For ix2 = lb2 To ub2: Let jagArr(ix1)(ix2) = arr2D(ix1, ix2): Next
        Next
    End If
    
Ending:
    Arr2DToJagArr = jagArr
End Function

''' @param arr() As Variant
''' @return As Object Is Scripting.Dictionary
Public Function CreateDictionary(ParamArray arr() As Variant) As Object
    Dim alen As Long: alen = UBound(arr)
    If Abs(alen Mod 2) = 0 Then Err.Raise 5
    
    Set CreateDictionary = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 0 To alen Step 2: CreateDictionary.Add arr(i), arr(i + 1): Next
End Function

''' @return As Object Is WScript.Shell
Public Property Get Wsh() As Object
    If xxWsh Is Nothing Then Set xxWsh = CreateObject("WScript.Shell")
    Set Wsh = xxWsh
End Property

''' @return As Object Is WbemScripting.SWbemLocator
Public Property Get Wmi() As Object
    If xxWmi Is Nothing Then Set xxWmi = CreateObject("WbemScripting.SWbemLocator")
    Set Wmi = xxWmi
End Property

''' @param ptrnFind As String
''' @param regexpOption As String
''' @return As Object Is VBScript.RegExp
Public Function CreateRegExp( _
    ByVal ptrnFind As String, Optional ByVal regexpOption As String = "" _
    ) As Object
    
    Dim cnt As Long: cnt = 0
    Set CreateRegExp = CreateObject("VBScript.RegExp")
    CreateRegExp.Pattern = ptrnFind
    CreateRegExp.Global = WithIncrIf(InStr(regexpOption, "g") > 0, True, cnt)
    CreateRegExp.IgnoreCase = WithIncrIf(InStr(regexpOption, "i") > 0, True, cnt)
    If cnt <> Len(regexpOption) Then Err.Raise 5
End Function
Private Function WithIncrIf( _
    ByVal expr As Variant, ByVal incif As Variant, ByRef cntr As Long _
    ) As Variant
    
    If Equals(expr, incif) Then cntr = cntr + 1
    WithIncrIf = expr
End Function

''' @param expr As String
''' @param ptrnFind As String
''' @param iCase As Boolean
''' @return As Variant(Of Array(Of String))
Public Function ReMatch( _
    ByVal expr As String, ByVal ptrnFind As String, _
    Optional ByVal iCase As Boolean = False _
    ) As Variant
    
    Dim ret As Variant: ret = Array()
    
    Dim regex As Object:  Set regex = CreateRegExp(ptrnFind, IIf(iCase, "i", ""))
    Dim ms As Object:     Set ms = regex.Execute(expr)
    If ms.Count < 1 Then: GoTo Ending
    
    Dim sms As Object:    Set sms = ms(0).SubMatches
    ReDim ret(sms.Count)
    
    ret(0) = ms.Item(0).Value
    Dim i As Integer
    For i = 1 To UBound(ret): ret(i) = sms.Item(i - 1): Next
    
Ending:
    ReMatch = ret
End Function

''' @param expr As String
''' @param ptrnFind As String
''' @param iCase As Boolean
''' @return As Variant(Of Array(Of Array(Of String)))
Public Function ReMatcheGlobal( _
    ByVal expr As String, ByVal ptrnFind As String, _
    Optional ByVal iCase As Boolean = False _
    ) As Variant
    
    Dim ret As Variant: ret = Array()
    
    Dim regex As Object: Set regex = CreateRegExp(ptrnFind, IIf(iCase, "i", "") & "g")
    Dim ms As Object:    Set ms = regex.Execute(expr)
    If ms.Count < 1 Then GoTo Ending
    
    ReDim ret(ms.Count - 1)
    
    Dim arr As Variant: ReDim arr(ms(0).SubMatches.Count)
    
    Dim i As Integer, j As Integer
    For i = 0 To UBound(ret)
        ret(i) = arr
        
        ret(i)(0) = ms.Item(i).Value
        For j = 1 To UBound(arr): ret(i)(j) = ms(i).SubMatches.Item(j - 1): Next
    Next
    
Ending:
    ReMatcheGlobal = ret
End Function

''' @param expr As String
''' @param ptrnFind As String
''' @param ptrnReplace As String
''' @param regexpOption As String
''' @return As Variant(Of Array(Of String))
Public Function ReReplace( _
    ByVal expr As String, ByVal ptrnFind As String, ByVal ptrnReplace As String, _
    Optional ByVal regexpOption As String = "" _
    ) As String
    
    Dim regex As Object: Set regex = CreateRegExp(ptrnFind, regexpOption)
    ReReplace = regex.Replace(expr, ptrnReplace)
End Function

''' @param expr As String
''' @param ptrnFind As String
''' @param iCase As Boolean
''' @return As String
Public Function ReTrim( _
    ByVal expr As String, ByVal ptrnFind As String, _
    Optional ByVal iCase As Boolean = False _
    ) As String
    
    ptrnFind = "^(?:" & ptrnFind & ")+|(?:" & ptrnFind & ")+$"
    
    Dim regex As Object: Set regex = CreateRegExp(ptrnFind, IIf(iCase, "i", "") & "g")
    ReTrim = regex.Replace(expr, "")
End Function

Private Function EvalScript(ByVal expr As String, ByVal lang As String) As String
    Dim doc As Object: Set doc = CreateObject("HtmlFile")
    doc.parentWindow.execScript "document.write(" & expr & ")", lang
    If Not doc.body Is Nothing Then EvalScript = doc.body.innerHTML
End Function

''' @param vbsExpr As String
''' @return As Variant
Public Function EvalVBS(ByVal vbsExpr As String) As String
    EvalVBS = EvalScript(vbsExpr, "VBScript")
End Function

''' @param jsExpr As String
''' @return As Variant
Public Function EvalJS(ByVal jsExpr As String) As String
    EvalJS = EvalScript(jsExpr, "JScript")
End Function

''' @return As Object Is StdRegProv
Public Function CreateStdRegProv() As Object
    Dim wmiSrv As Object: Set wmiSrv = Wmi.ConnectServer(, "root\default")
    Set CreateStdRegProv = wmiSrv.Get("StdRegProv")
End Function
