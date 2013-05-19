Attribute VB_Name = "Core"
Option Explicit
Option Private Module

''' @return VARIANT
''' @param [in] IDispatch* Object
''' @param [in] BSTR ProcName
''' @param [in] VbCallType CallType
''' @param [in] SAFEARRAY(VARIANT)* Args
''' @param [in, lcid] long lcid
#If VBA7 Then
#If Win64 Then
Public Declare PtrSafe _
Function rtcCallByName Lib "VBE7.DLL" ( _
    ByVal Object As Object, _
    ByVal ProcName As LongPtr, _
    ByVal CallType As VbCallType, _
    ByRef Args() As Any _
    ) As Variant
#Else
Public Declare _
Function rtcCallByName Lib "VBE7.DLL" ( _
    ByVal Object As Object, _
    ByVal ProcName As Long, _
    ByVal CallType As VbCallType, _
    ByRef Args() As Any _
    ) As Variant
#End If
#Else
Public Declare _
Function rtcCallByName Lib "VBE6.DLL" ( _
    ByVal Object As Object, _
    ByVal ProcName As Long, _
    ByVal CallType As VbCallType, _
    ByRef Args() As Any _
    ) As Variant
#End If

Public Property Get Missing() As Variant
    Missing = GetMissing()
End Property

Private Function GetMissing(Optional ByVal mss As Variant) As Variant
    'If Not IsMissing(mss) Then Err.Raise 5
    GetMissing = mss
End Function

''' @usage
'''     Dim i as Integer: i = 42
'''     IncrPre(i)  '43
'''     i           '43
''' @param n As Long
''' @return As Long
Public Function IncrPre(ByRef n As Variant, Optional ByVal stepVal As Variant = 1) As Variant
    n = n + stepVal: IncrPre = n
End Function

''' @usage
'''     Dim i as Integer: i = 42
'''     IncrPst(i)  '42
'''     i           '43
''' @param n As Long
''' @return As Long
Public Function IncrPst(ByRef n As Variant, Optional ByVal stepVal As Variant = 1) As Variant
    IncrPst = n: n = n + stepVal
End Function

Public Function ToStr(ByVal x As Variant) As String
    If IsObject(x) Then
        On Error Resume Next
        ToStr = TypeName(x)
        ToStr = x.ToStr()
    Else
        ToStr = x
    End If
End Function

''' @usage
'''     Init(New Tuple2, "A", 4) 'Tuple2 { Item1 = "A", Item2 = 4 }
''' @param obj As Object Is T
''' @param args As Variant()
''' @return As Object Is T
Public Function Init(ByVal obj As Object, ParamArray params() As Variant) As Object
    Dim i As Long
    Dim ubParam As Long: ubParam = UBound(params)
    Dim ps() As Variant: ReDim ps(ubParam)
    For i = 0 To ubParam: ps(i) = params(i): Next
    
    rtcCallByName obj, StrPtr("Init"), VbMethod, ps
    
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
'''     Max(3, 6, 5) '6
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
        ObjArrMSort arr, ix0
    Else
        ValArrMSort arr, ix0
    End If
    
Escape:
End Sub
Private Sub ObjArrMSort(arr As Variant, lb As Long)
    Dim alen As Long: alen = ArrLen(arr)
    If alen <= 1 Then GoTo Escape
    
    '' optimization
    If alen <= 8 Then
        ObjArrISort arr, lb
        GoTo Escape
    End If
    
    Dim i As Long
    Dim l1 As Long: l1 = Fix(alen / 2)
    Dim l2 As Long: l2 = alen - l1
    
    Dim ub1 As Long:   ub1 = lb + l1 - 1
    Dim a1 As Variant: ReDim a1(lb To ub1)
    For i = lb To ub1: Set a1(i) = arr(i): Next
    ObjArrMSort a1, lb
    
    Dim ub2 As Long:   ub2 = lb + l2 - 1
    Dim a2 As Variant: ReDim a2(lb To ub2)
    For i = lb To ub2: Set a2(i) = arr(l1 + i): Next
    ObjArrMSort a2, lb
    
    Dim i1 As Long: i1 = lb
    Dim i2 As Long: i2 = lb
    While i1 <= ub1 Or i2 <= ub2
        If ArrMergeSw(a1, i1, ub1, a2, i2, ub2) Then
            Set arr(i1 + i2 - lb) = a1(IncrPst(i1))
        Else
            Set arr(i1 + i2 - lb) = a2(IncrPst(i2))
        End If
    Wend
    
Escape:
End Sub
Private Sub ValArrMSort(arr As Variant, lb As Long)
    Dim alen As Long: alen = ArrLen(arr)
    If alen <= 1 Then GoTo Escape
    
    '' optimization
    If alen <= 8 Then
        ValArrISort arr, lb
        GoTo Escape
    End If
    
    Dim i As Long
    Dim l1 As Long: l1 = Fix(alen / 2)
    Dim l2 As Long: l2 = alen - l1
    
    Dim ub1 As Long:   ub1 = lb + l1 - 1
    Dim a1 As Variant: ReDim a1(lb To ub1)
    For i = lb To ub1: Let a1(i) = arr(i): Next
    ValArrMSort a1, lb
    
    Dim ub2 As Long:   ub2 = lb + l2 - 1
    Dim a2 As Variant: ReDim a2(lb To ub2)
    For i = lb To ub2: Let a2(i) = arr(l1 + i): Next
    ValArrMSort a2, lb
    
    Dim i1 As Long: i1 = lb
    Dim i2 As Long: i2 = lb
    While i1 <= ub1 Or i2 <= ub2
        If ArrMergeSw(a1, i1, ub1, a2, i2, ub2) Then
            Let arr(i1 + i2 - lb) = a1(IncrPst(i1))
        Else
            Let arr(i1 + i2 - lb) = a2(IncrPst(i2))
        End If
    Wend
    
Escape:
End Sub
Private Sub ObjArrISort(arr As Variant, lb As Long)
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
Private Sub ValArrISort(arr As Variant, lb As Long)
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

''' @usage
'''     ArrConcat(Array(1, 2, 3), Array(4, 5)) ' Array(1, 2, 3, 4, 5)
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

''' @usage
'''     ArrFlatten(Array(Array(1, 2), Array(3), Array(4, 5))) ' Array(1, 2, 3, 4, 5)
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

''' @usage
'''     ArrRange(1, 9) 'Array(1, 2, 3, 4, 5, 6, 7, 8, 9)
''' @param fromVal As Variant(Of T)
''' @param toVal As Variant(Of T)
''' @param stepVal As Variant(Of T)
''' @return As Variant(Of Array(Of T))
Public Function ArrRange( _
    ByVal fromVal As Variant, ByVal toVal As Variant, Optional ByVal stepVal As Variant = 1 _
    ) As Variant
    
    'FIXME: parameters type check
    
    Dim i As Long: i = 0
    Dim alen As Long: alen = 32
    Dim arr As Variant: ReDim arr(alen - 1)
    
    Select Case stepVal
    Case Is > 0
        Do While fromVal <= toVal
            arr(IncrPst(i)) = IncrPst(fromVal, stepVal)
            If i >= alen Then alen = alen * 2: ReDim Preserve arr(alen - 1)
        Loop
    Case Is < 0
        Do While fromVal >= toVal
            arr(IncrPst(i)) = IncrPst(fromVal, stepVal)
            If i >= alen Then alen = alen * 2: ReDim Preserve arr(alen - 1)
        Loop
    Case Else
        Err.Raise 5
    End Select
    
    If i > 0 Then
        ReDim Preserve arr(i - 1)
    Else
        arr = Array()
    End If
    ArrRange = arr
End Function

''' @usage
'''     ' Function Twice(ByVal n As Integer) As Integer
'''     ArrMap(Init(New Func, AddressOf Twice, vbInteger), ArrRange(1, 9))
'''     ' => Array(2, 4, 6, 8, 10, 12, 14, 16, 18)
''' @param f As Func(Of T, U)
''' @param arr As Variant(Of Array(Of T))
''' @return As Variant(Of Array(Of U))
Public Function ArrMap(ByVal f As Func, ByVal arr As Variant) As Variant
    If Not IsArray(arr) Then Err.Raise 13
    Dim lb As Long: lb = LBound(arr)
    Dim ub As Long: ub = UBound(arr)
    Dim ret As Variant
    If ub - lb < 0 Then
        ret = Array()
        GoTo Ending
    End If
    
    ReDim ret(lb To ub)
    
    Dim i As Long
    For i = lb To ub: f.FastApply ret(i), arr(i): Next
    
Ending:
    ArrMap = ret
End Function

''' @usage
'''     ' Function IsOdd(ByVal n As Integer) As Boolean
'''     ArrFilter(Init(New Func, vbBoolean, AddressOf IsOdd), ArrRange(1, 9))
'''     ' => Array(1, 3, 5, 7, 9)
''' @param f As Func(Of T, Boolean)
''' @param arr As Variant(Of Array(Of T))
''' @return As Variant(Of Array(Of T))
Public Function ArrFilter(ByVal f As Func, ByVal arr As Variant) As Variant
    If Not IsArray(arr) Then Err.Raise 13
    Dim lb As Long: lb = LBound(arr)
    Dim ub As Long: ub = UBound(arr)
    Dim ret As Variant
    If ub - lb < 0 Then
        ret = Array()
        GoTo Ending
    End If
    
    ReDim ret(lb To ub)
    
    Dim flg As Boolean
    Dim ixArr As Long
    Dim ixRet As Long: ixRet = lb
    If IsObject(arr(lb)) Then
        For ixArr = lb To ub
            f.FastApply flg, arr(ixArr)
            If flg Then Set ret(IncrPst(ixRet)) = arr(ixArr)
        Next
    Else
        For ixArr = lb To ub
            f.FastApply flg, arr(ixArr)
            If flg Then Let ret(IncrPst(ixRet)) = arr(ixArr)
        Next
    End If
    
    If ixRet > 0 Then
        ReDim Preserve ret(lb To ixRet - 1)
    Else
        ret = Array()
    End If
    
Ending:
    ArrFilter = ret
End Function

''' @usage
'''     ' Function Add(ByVal i As Integer, ByVal j As Integer) As Integer
'''     ArrFold(Init(New Func, AddressOf Add, vbInteger), ArrRange(1, 100), 0) '5050
''' @param f As Func(Of U, T, U)
''' @param arr As Variant(Of Array(Of T))
''' @param seedVal As Variant(Of U)
''' @return As Variant(Of U)
Public Function ArrFold(ByVal f As Func, ByVal arr As Variant, Optional ByVal seedVal As Variant _
    ) As Variant
    
    If Not IsArray(arr) Then Err.Raise 13
    
    Dim stat As Variant
    Dim i As Long: i = LBound(arr)
    If IsMissing(seedVal) Then
        stat = arr(IncrPst(i))
    Else
        stat = seedVal
    End If
    
    For i = i To UBound(arr): f.FastApply stat, stat, arr(i): Next
    
    If IsObject(stat) Then
        Set ArrFold = stat
    Else
        Let ArrFold = stat
    End If
End Function

''' @usage
'''     ' Function FibFun(ByVal pair As Variant) As Variant
'''     ArrUnfold(Init(New Func, AddressOf FibFun, vbVariant, vbVariant), Array(1, 1))
'''     ' => 1, 2, 3, 5, 8, 13, 21, 34, 55, 89
''' @param f As Func
''' @param seedVal As Variant(Of T)
''' @return As Variant(Of Array(Of U))
Public Function ArrUnfold(ByVal f As Func, ByVal seedVal As Variant) As Variant
    Dim i As Long: i = 0
    Dim alen As Long: alen = 32
    Dim arr As Variant: ReDim arr(alen - 1)
    
    Dim stat As Variant
    If IsObject(seedVal) Then
        f.FastApply stat, seedVal
        Do Until IsMissing(stat(1))
            Set arr(IncrPst(i)) = stat(0)
            If i >= alen Then alen = alen * 2: ReDim Preserve arr(alen - 1)
            f.FastApply stat, stat(1)
        Loop
    Else
        f.FastApply stat, seedVal
        Do Until IsMissing(stat(1))
            Let arr(IncrPst(i)) = stat(0)
            If i >= alen Then alen = alen * 2: ReDim Preserve arr(alen - 1)
            f.FastApply stat, stat(1)
        Loop
    End If
    
    If i > 0 Then
        ReDim Preserve arr(i - 1)
    Else
        arr = Array()
    End If
    ArrUnfold = arr
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

''' @usage
'''     BitFlag(False, True)                '1
'''     BitFlag(True, False, False, False)  '8
'''     BitFlag(1, 0, 0, 0)                 '8
''' @param flgs() As Variant(Of Boolean)
''' @return As Long
Public Function BitFlag(ParamArray flgs() As Variant) As Long
    BitFlag = 0
    Dim ub As Long: ub = UBound(flgs)
    
    Dim i As Long
    For i = 0 To ub
        BitFlag = BitFlag + IIf(flgs(i), 1, 0) * 2 ^ (ub - i)
    Next
End Function

Public Function ARound( _
    ByVal num As Variant, Optional ByVal digits As Integer = 0, Optional rndup As Integer = 5 _
    ) As Variant
    
    If Not IsNumeric(num) Then Err.Raise 13
    If Not (1 <= rndup And rndup <= 10) Then Err.Raise 5
    
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

'''
''' following functions for japanese only.
'''

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
        ixByte = ixByte + SjisByteNum(Mid(str, IncrPst(ixStr), 1))
    Wend
    LeftA = Left(str, ixStr - (ixByte - byteLen))
End Function

Public Function RightA(ByVal str As String, ByVal byteLen As Long) As String
    Dim idxs As Object: Set idxs = CreateObject("Scripting.Dictionary")
    Dim ixByte As Long: ixByte = 1
    Dim ixStr As Long:  ixStr = 1
    While ixStr <= Len(str)
        idxs.Add ixByte, ixStr
        ixByte = ixByte + SjisByteNum(Mid(str, IncrPst(ixStr), 1))
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
        ixByte = ixByte + SjisByteNum(Mid(str, IncrPst(ixStr), 1))
    Wend
    
    Dim n As Long: n = ixStr - (ixByte - byteLen)
    SepA = Init(New Tuple2, Left(str, n), Mid(str, n + 1, strLen))
End Function
