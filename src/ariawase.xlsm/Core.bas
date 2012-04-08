Attribute VB_Name = "Core"
Option Explicit
Option Private Module

''' @usage
'''     Init(New Tuple2, "A", 4) 'Tuple2 { Item1 = "A", Item2 = 4 }
''' @param obj As Object Is T
''' @param args As Variant()
''' @return As Object Is T
Public Function Init(ByVal obj As Object, ParamArray args() As Variant) As Object
    Select Case UBound(args)
        Case -1: obj.Init
        Case 0:  obj.Init args(0)
        Case 1:  obj.Init args(0), args(1)
        Case 2:  obj.Init args(0), args(1), args(2)
        Case 3:  obj.Init args(0), args(1), args(2), args(3)
        Case 4:  obj.Init args(0), args(1), args(2), args(3), args(4)
        Case 5:  obj.Init args(0), args(1), args(2), args(3), args(4), args(5)
        Case 6:  obj.Init args(0), args(1), args(2), args(3), args(4), args(5), args(6)
        Case 7:  obj.Init args(0), args(1), args(2), args(3), args(4), args(5), args(6), args(7)
        Case Else: Err.Raise 5
    End Select
    Set Init = obj
End Function

''' @usage
'''     Eq(42, 42)              'True
'''     42 = "42"               'True
'''     CVar(42) = CVar("42")   'False
'''     Eq(42, "42")            'False
'''     CVar(Empty) = CVar(0) And Eq(Empty, 0)  'True
'''     Eq(Init(New Tuple2, "A", 4), Init(New Tuple2, "A", 4)) 'False
''' @param x As Variant(Of T)
''' @param y As Variant(Of T)
''' @return As Variant(Of Nullable(Of Boolean))
Public Function Eq(ByVal x As Variant, ByVal y As Variant) As Variant
    Dim xIsObj As Boolean: xIsObj = IsObject(x)
    Dim yIsObj As Boolean: yIsObj = IsObject(y)
    If xIsObj Xor yIsObj Then Err.Raise 13
    
    If xIsObj And yIsObj Then
        Eq = x Is y
    Else
        Eq = x = y 'Nullable
    End If
End Function

''' @usage
'''     Equals(Empty, Empty)    'True
'''     Equals(Empty, 0)        'Err.Raise 13
'''     Equals(0, 0.0)          'Err.Raise 13
'''     Equals("", vbNullString)    'True
'''     Equals(Init(New Tuple2, "A", 4), Init(New Tuple2, "A", 4)) 'True
''' @param x As Variant(Of T)
''' @param y As Variant(Of T)
''' @return As Variant(Of Nullable(Of Boolean))
Public Function Equals(ByVal x As Variant, ByVal y As Variant) As Variant
    Dim xIsObj As Boolean: xIsObj = IsObject(x)
    Dim yIsObj As Boolean: yIsObj = IsObject(y)
    If xIsObj Xor yIsObj Then Err.Raise 13
    
    If xIsObj And yIsObj Then
        Equals = x.Equals(y)
    Else
        If TypeName(x) = TypeName(y) Then
            Equals = x = y 'Nullable
        Else
            Err.Raise 13
        End If
    End If
    Exit Function
End Function

''' @usage
'''     Compare(3, 9) '-1
'''     Compare(5, 5) ' 0
'''     Compare(9, 3) ' 1
'''     Compare(Init(New Tuple2, 2, ".txt"), Init(New Tuple2, 10, ".txt"))      '-1
'''     Compare(Init(New Tuple2, "2", ".txt"), Init(New Tuple2, "10", ".txt"))  ' 1
''' @param x As Variant(Of T)
''' @param y As Variant(Of T)
''' @return As Variant(Of Nullable(Of Integer))
Public Function Compare(ByVal x As Variant, ByVal y As Variant) As Variant
    Dim xIsObj As Boolean: xIsObj = IsObject(x)
    Dim yIsObj As Boolean: yIsObj = IsObject(y)
    If xIsObj Xor yIsObj Then Err.Raise 13
    
    If xIsObj And yIsObj Then
        Compare = x.Compare(y)
    Else
        If TypeName(x) <> TypeName(y) Then Err.Raise 13
        
        If x = y Then Compare = 0 Else _
        If x < y Then Compare = -1 Else _
        If x > y Then Compare = 1 Else _
        Compare = Null
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

''' @usage
'''     Min(3, 6, 5) '3
''' @param arr() As Variant(Of T)
''' @return As Variant(Of T)
Public Function Min(ParamArray arr() As Variant) As Variant
    MinOrMax arr, -1, Min
End Function

''' @usage
'''     Min(3, 6, 5) '6
''' @param arr() As Variant(Of T)
''' @return As Variant(Of T)
Public Function Max(ParamArray arr() As Variant) As Variant
    MinOrMax arr, 1, Max
End Function

''' @usage
'''     ArrLen(Array("V", "B", "A")) '3
''' @param arr As Variant(Of Array(Of T))
''' @return As Long
Public Function ArrLen(ByVal arr As Variant, Optional ByVal dimen As Integer = 1) As Long
    If Not IsArray(arr) Then Err.Raise 13
    ArrLen = UBound(arr, dimen) - LBound(arr, dimen) + 1
End Function

''' @usage
'''     ArrEquals(Array(0, 1, 2), Array(0, 1, 2)) 'True
'''     ArrEquals(Array(0, 1, 2), Array(2, 1, 0)) 'False
''' @param arr1 As Variant(Of Array(Of T))
''' @param arr2 As Variant(Of Array(Of T))
''' @return As Boolean
Public Function ArrEquals(ByVal arr1 As Variant, ByVal arr2 As Variant) As Boolean
    If Not (IsArray(arr1) And IsArray(arr2)) Then Err.Raise 13
    ArrEquals = False
    
    Dim alen As Long: alen = ArrLen(arr1)
    If alen <> ArrLen(arr2) Then GoTo Escape
    
    Dim ix0 As Long: ix0 = LBound(arr1)
    Dim pad As Long: pad = LBound(arr2) - ix0
    
    Dim i As Long
    For i = ix0 To ix0 + alen - 1
        If Not Equals(arr1(i), arr2(pad + i)) Then GoTo Escape
    Next
    ArrEquals = True
    
Escape:
End Function

''' @usage
'''     ArrIndexOf(Array("V", "B", "A"), "A")       ' 2
'''     ArrIndexOf(Array("V", "B", "A"), "Z")       '-1
'''     ArrIndexOf(Array("I", "I", "f"), "I", 1)    ' 1
'''     ArrIndexOf(Array("I", "I", "f"), "f", 0, 2) '-1
'''     ArrIndexOf(Array("I", "I", "f"), "f", 1, 9) ' 2
''' @param arr As Variant(Of Array(Of T))
''' @param val As Variant(Of T)
''' @param ixStart As Variant(Of Empty Or Long)
''' @param cnt As Long
''' @return As Long
Public Function ArrIndexOf( _
    ByVal arr As Variant, ByVal val As Variant, _
    Optional ByVal ixStart As Variant = Empty, Optional ByVal cnt As Long = -1 _
    ) As Long
    
    If Not IsArray(arr) Then Err.Raise 13
    
    Dim ix0 As Long: ix0 = LBound(arr)
    If IsEmpty(ixStart) Then ixStart = ix0
    If IsNumeric(ixStart) Then ixStart = CLng(ixStart) Else Err.Raise 13
    If ixStart < ix0 Then Err.Raise 5
    If cnt < 0 Then cnt = ArrLen(arr) Else cnt = Min(cnt, ArrLen(arr))
    
    ArrIndexOf = ixStart - 1
    
    Dim i As Long
    For i = ixStart To ixStart + cnt - 1
        If Equals(arr(i), val) Then
            ArrIndexOf = i
            GoTo Escape
        End If
    Next
    
Escape:
End Function

''' @usage
'''     Dim arr As Variant: arr = Array("S", "O", "R", "T")
'''     ArrSort arr
'''     arr 'Array("O", "R", "S", "T")
''' @param arr As Variant(Of Array(Of T))
Public Sub ArrSort(ByRef arr As Variant)
    If Not IsArray(arr) Then Err.Raise 13
    If ArrLen(arr) <= 1 Then GoTo Escape
    
    Dim ix0 As Long: ix0 = LBound(arr)
    If IsObject(arr(ix0)) Then
        ObjArrSort arr, ix0
    Else
        ValArrSort arr, ix0
    End If
    
Escape:
End Sub
Private Sub ObjArrSort(arr As Variant, lb As Long)
    Dim alen As Long: alen = ArrLen(arr)
    If alen <= 1 Then GoTo Escape
    
    If alen <= 8 Then
        ObjArrSortI arr, lb
        GoTo Escape
    End If
    
    Dim i As Long
    Dim l1 As Long: l1 = Fix(alen / 2)
    Dim l2 As Long: l2 = alen - l1
    
    Dim ub1 As Long:   ub1 = lb + l1 - 1
    Dim a1 As Variant: ReDim a1(lb To ub1)
    For i = lb To ub1: Set a1(i) = arr(i): Next
    ObjArrSort a1, lb
    
    Dim ub2 As Long:   ub2 = lb + l2 - 1
    Dim a2 As Variant: ReDim a2(lb To ub2)
    For i = lb To ub2: Set a2(i) = arr(l1 + i): Next
    ObjArrSort a2, lb
    
    Dim i1 As Long: i1 = lb
    Dim i2 As Long: i2 = lb
    While i1 <= ub1 Or i2 <= ub2
        If ArrMergeSw(a1, i1, ub1, a2, i2, ub2) Then
            Set arr(i1 + i2 - lb) = a1(i1): i1 = i1 + 1
        Else
            Set arr(i1 + i2 - lb) = a2(i2): i2 = i2 + 1
        End If
    Wend
    
Escape:
End Sub
Private Sub ValArrSort(arr As Variant, lb As Long)
    Dim alen As Long: alen = ArrLen(arr)
    If alen <= 1 Then GoTo Escape
    
    If alen <= 8 Then
        ValArrSortI arr, lb
        GoTo Escape
    End If
    
    Dim i As Long
    Dim l1 As Long: l1 = Fix(alen / 2)
    Dim l2 As Long: l2 = alen - l1
    
    Dim ub1 As Long:   ub1 = lb + l1 - 1
    Dim a1 As Variant: ReDim a1(lb To ub1)
    For i = lb To ub1: Let a1(i) = arr(i): Next
    ValArrSort a1, lb
    
    Dim ub2 As Long:   ub2 = lb + l2 - 1
    Dim a2 As Variant: ReDim a2(lb To ub2)
    For i = lb To ub2: Let a2(i) = arr(l1 + i): Next
    ValArrSort a2, lb
    
    Dim i1 As Long: i1 = lb
    Dim i2 As Long: i2 = lb
    While i1 <= ub1 Or i2 <= ub2
        If ArrMergeSw(a1, i1, ub1, a2, i2, ub2) Then
            Let arr(i1 + i2 - lb) = a1(i1): i1 = i1 + 1
        Else
            Let arr(i1 + i2 - lb) = a2(i2): i2 = i2 + 1
        End If
    Wend
    
Escape:
End Sub
Private Sub ObjArrSortI(arr As Variant, lb As Long)
    Dim i As Long, j As Long, x As Variant
    For i = lb + 1 To UBound(arr)
        j = i
        Do While j > lb
            If Compare(arr(j - 1), arr(j)) <= 0 Then Exit Do
            Set x = arr(j): Set arr(j) = arr(j - 1): Set arr(j - 1) = x
            j = j - 1
        Loop
    Next
End Sub
Private Sub ValArrSortI(arr As Variant, lb As Long)
    Dim i As Long, j As Long, x As Variant
    For i = lb + 1 To UBound(arr)
        j = i
        Do While j > lb
            If Compare(arr(j - 1), arr(j)) <= 0 Then Exit Do
            Let x = arr(j): Let arr(j) = arr(j - 1): Let arr(j - 1) = x
            j = j - 1
        Loop
    Next
End Sub
Private Function ArrMergeSw( _
    arr1 As Variant, i1 As Long, ub1 As Long, _
    arr2 As Variant, i2 As Long, ub2 As Long _
    ) As Boolean
    
    If i1 > ub1 Then ArrMergeSw = False Else _
    If i2 > ub2 Then ArrMergeSw = True Else _
    ArrMergeSw = Compare(arr1(i1), arr2(i2)) < 1
End Function

''' @usage
'''     ArrUniq(Array(6, 5, 5, 3, 6)) ' Array(6, 5, 3)
''' @param arr As Variant(Of Array(Of T))
Public Function ArrUniq(ByVal arr As Variant) As Variant
    If Not IsArray(arr) Then Err.Raise 13
    Dim ret As Variant: ret = Array()
    Dim lbA As Long: lbA = LBound(arr)
    Dim ubA As Long: ubA = UBound(arr)
    If ubA - lbA < 0 Then GoTo Ending
    
    ReDim ret(lbA To ubA)
    
    Dim ixA As Long, ixR As Long: ixR = lbA
    Dim isObj As Boolean: isObj = IsObject(arr(lbA))
    If isObj Then
        For ixA = lbA To ubA
            If ArrIndexOf(ret, arr(ixA), lbA, ixR - lbA) < lbA Then
                Set ret(ixR) = arr(ixA)
                ixR = ixR + 1
            End If
        Next
    Else
        For ixA = lbA To ubA
            If ArrIndexOf(ret, arr(ixA), lbA, ixR - lbA) < lbA Then
                Let ret(ixR) = arr(ixA)
                ixR = ixR + 1
            End If
        Next
    End If
    
    ReDim Preserve ret(lbA To ixR - 1)
    
Ending:
    ArrUniq = ret
End Function

''' @param arr As Variant(Of Array(Of T))
''' @return As Collection(Of T)
Public Function ArrToClct(ByVal arr As Variant) As Collection
    If Not IsArray(arr) Then Err.Raise 13
    Set ArrToClct = New Collection
    Dim v As Variant
    For Each v In arr: ArrToClct.Add v: Next
End Function

''' @param clct as As Collection(Of T)
''' @return As Variant(Of Array(Of T))
Public Function ClctToArr(ByVal clct As Collection) As Variant
    Dim arr As Variant: arr = Array()
    Dim clen As Long: clen = clct.Count
    If clen < 1 Then GoTo Ending
    
    ReDim arr(clen - 1)
    Dim i As Long: i = 0
    Dim v As Variant
    If IsObject(clct.Item(1)) Then
        For Each v In clct: Set arr(i) = v: i = i + 1: Next
    Else
        For Each v In clct: Let arr(i) = v: i = i + 1: Next
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

Public Function CreateAssocArray(ParamArray arr() As Variant) As Variant
    Dim alen As Long: alen = UBound(arr)
    If Abs(alen Mod 2) = 0 Then Err.Raise 5
    
    Dim aarr As Variant: aarr = Array()
    If alen < 0 Then GoTo Ending
    
    ReDim aarr(Fix(UBound(arr) / 2))
    Dim i As Long
    For i = 0 To UBound(aarr): Set aarr(i) = Init(New Tuple2, arr(2 * i), arr(2 * i + 1)): Next
    
Ending:
    CreateAssocArray = aarr
End Function

Public Function CreateDictionary(ParamArray arr() As Variant) As Object
    Dim alen As Long: alen = UBound(arr)
    If Abs(alen Mod 2) = 0 Then Err.Raise 5
    
    Set CreateDictionary = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 0 To alen Step 2: CreateDictionary.Add arr(i), arr(i + 1): Next
End Function

Public Function AssocArrToDict(ByVal aarr As Variant) As Object
    If Not IsArray(aarr) Then Err.Raise 13
    Set AssocArrToDict = CreateDictionary()
    Dim v As Variant '(Of Tuple2)
    For Each v In aarr: AssocArrToDict.Add v.Item1, v.Item2: Next
End Function

Public Function DictToAssocArr(ByVal dict As Object) As Variant
    If TypeName(dict) <> "Dictionary" Then Err.Raise 13
    Dim arr As Variant: arr = Array()
    
    Dim ks As Variant: ks = dict.Keys
    Dim dlen As Long: dlen = UBound(ks)
    If dlen < 0 Then GoTo Ending
    
    ReDim arr(UBound(ks))
    Dim i As Long
    For i = 0 To dlen: Set arr(i) = Init(New Tuple2, ks(i), dict.Item(ks(i))): Next
    
Ending:
    DictToAssocArr = arr
End Function

Public Function ARound( _
    ByVal num As Variant, Optional ByVal digits As Integer = 0, Optional rndup As Integer = 5 _
    ) As Variant
    
    If Not IsNumeric(num) Then Err.Raise 13
    If Not (1 <= rndup And rndup <= 9) Then Err.Raise 5
    
    Dim n As Integer: n = CDec(10 ^ Abs(digits))
    Dim z As Double:  z = CDec(Sgn(num) * 0.1 * (10 - rndup))
    If digits >= 0 Then
        ARound = Fix(num * n + z) / n
    Else
        ARound = Fix(num / n + z) * n
    End If
    
Escape:
End Function

Public Function BeginOfMonth(ByVal dt As Date) As Date
    BeginOfMonth = DateAdd("d", -Day(dt) + 1, dt)
End Function

Public Function EndOfMonth(ByVal dt As Date) As Date
    EndOfMonth = DateAdd("d", -1, BeginOfMonth(DateAdd("m", 1, dt)))
End Function

Public Function BeginOfWeek(ByVal dt As Date, Optional fstDayOfWeek As VbDayOfWeek = vbSunday) As Date
    BeginOfWeek = DateAdd("d", 1 - Weekday(dt, fstDayOfWeek), dt)
End Function

Public Function EndOfWeek(ByVal dt As Date, Optional fstDayOfWeek As VbDayOfWeek = vbSunday) As Date
    EndOfWeek = DateAdd("d", 7 - Weekday(dt, fstDayOfWeek), dt)
End Function

''' @param str As String Is Char
''' @return As Integer
Private Function SjisByteNum(ByVal str As String) As Integer
   Dim x As Integer
   x = Asc(str) / &H100 And &HFF
   SjisByteNum = IIf((&H81 <= x And x <= &H9F) Or (&HE0 <= x And x <= &HFC), 2, 1)
End Function

Public Function LeftA(ByVal str As String, ByVal byteLen As Long) As String
    Dim ixByte As Long: ixByte = 1
    Dim ixStr As Long:  ixStr = 1
    While (ixByte < 1 + byteLen) And (ixStr <= Len(str))
        ixByte = ixByte + SjisByteNum(Mid(str, ixStr, 1))
        ixStr = ixStr + 1
    Wend
    LeftA = Left(str, ixStr - (ixByte - byteLen))
End Function

Public Function RightA(ByVal str As String, ByVal byteLen As Long) As String
    Dim idxs As Object: Set idxs = CreateObject("Scripting.Dictionary")
    Dim ixByte As Long: ixByte = 1
    Dim ixStr As Long:  ixStr = 1
    While ixStr <= Len(str)
        idxs.Add ixByte, ixStr
        ixByte = ixByte + SjisByteNum(Mid(str, ixStr, 1))
        ixStr = ixStr + 1
    Wend
    idxs.Add ixByte, ixStr
    
    For byteLen = byteLen To 0 Step -1
        If idxs.Exists(ixByte - byteLen) Then Exit For
    Next
    
    RightA = Right(str, ixStr - idxs.Item(ixByte - byteLen))
End Function

Public Function SepA(ByVal str As String, ByVal byteLen As Long) As Tuple2
    Dim ixByte As Long: ixByte = 1
    Dim ixStr  As Long: ixStr = 1
    Dim strLen As Long: strLen = Len(str)
    
    While (ixByte < 1 + byteLen) And (ixStr <= strLen)
        ixByte = ixByte + SjisByteNum(Mid(str, ixStr, 1))
        ixStr = ixStr + 1
    Wend
    
    Dim n As Long: n = ixStr - (ixByte - byteLen)
    SepA = Init(New Tuple2, Left(str, n), Mid(str, n + 1, strLen))
End Function
