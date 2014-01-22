Attribute VB_Name = "Core"
''' +----                                                             --+ '''
''' |                          Ariawase 0.5.0                           | '''
''' |             Ariawase is free library for VBA cowboys.             | '''
''' |        The Project Page: https://github.com/igeta/Ariawase        | '''
''' +--                                                             ----+ '''
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
    ByRef Args() As Any, _
    Optional ByVal lcid As Long _
    ) As Variant
#Else
Public Declare _
Function rtcCallByName Lib "VBE7.DLL" ( _
    ByVal Object As Object, _
    ByVal ProcName As Long, _
    ByVal CallType As VbCallType, _
    ByRef Args() As Any, _
    Optional ByVal lcid As Long _
    ) As Variant
#End If
#Else
Public Declare _
Function rtcCallByName Lib "VBE6.DLL" ( _
    ByVal Object As Object, _
    ByVal ProcName As Long, _
    ByVal CallType As VbCallType, _
    ByRef Args() As Any, _
    Optional ByVal lcid As Long _
    ) As Variant
#End If

Public Property Get Missing() As Variant
    Missing = GetMissing()
End Property

Private Function GetMissing(Optional ByVal mss As Variant) As Variant
    'If Not IsMissing(mss) Then Err.Raise 5
    GetMissing = mss
End Function

''' @param n As Long
''' @return As Long
Public Function IncrPre(ByRef n As Variant, Optional ByVal stepVal As Variant = 1) As Variant
    n = n + stepVal: IncrPre = n
End Function

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

Public Function ToLiteral(ByVal x As Variant) As String
    Dim ty As String: ty = TypeName(x)
    Select Case ty
        Case "Byte":        ToLiteral = "CByte(" & x & ")"
        Case "Integer":     ToLiteral = x & "%"
        Case "Long":        ToLiteral = x & "&"
        Case "Single":      ToLiteral = x & "!"
        Case "Double":      ToLiteral = x & "#"
        Case "Currency":    ToLiteral = x & "@"
        Case "Decimal":     ToLiteral = "CDec(" & x & ")"
        Case "Date":        ToLiteral = "#" & x & "#"
        Case "Boolean":     ToLiteral = x
        Case "Empty":       ToLiteral = "(Empty)"
        Case "Null":        ToLiteral = "(Null)"
        Case "Nothing":     ToLiteral = "(Nothing)"
        Case "Unknown":     ToLiteral = "(Unknown)"
        Case "Error":       ToLiteral = "(Error)"
        Case "ErrObject":   ToLiteral = "Err " & x.Number
        Case "String"
            If StrPtr(x) = 0 Then
                ToLiteral = "(vbNullString)"
            Else
                ToLiteral = """" & x & """"
            End If
        Case Else
            If Right(ty, 2) = "()" Then
                '' FIXME: for multidimensional array
                Dim i As Long
                For i = LBound(x) To UBound(x): x(i) = ToLiteral(x(i)): Next
                ToLiteral = "Array(" & Join(x, ", ") & ")"
            Else
                On Error Resume Next
                ToLiteral = ty
                ToLiteral = x.ToStr()
            End If
    End Select
End Function

''' @param obj As Object Is T
''' @param args As Variant()
''' @return As Object Is T
Public Function Init(ByVal obj As Object, ParamArray params() As Variant) As Object
    Dim i As Long
    Dim ubParam As Long: ubParam = UBound(params)
    Dim ps() As Variant: ReDim ps(ubParam)
    For i = 0 To ubParam
        If IsObject(params(i)) Then
            Set ps(i) = params(i)
        Else
            Let ps(i) = params(i)
        End If
    Next
    rtcCallByName obj, StrPtr("Init"), VbMethod, ps
    
    Set Init = obj
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
''' @return As Variant(Of Boolean Or Null Or Empty)
Public Function Equals(ByVal x As Variant, ByVal y As Variant) As Variant
    Dim xIsObj As Boolean: xIsObj = IsObject(x)
    Dim yIsObj As Boolean: yIsObj = IsObject(y)
    If xIsObj Xor yIsObj Then
        Equals = Empty
    ElseIf xIsObj And yIsObj Then
        Equals = x.Equals(y)
    Else
        If TypeName(x) = TypeName(y) Then
            Equals = x = y
        ElseIf IsNull(x) Or IsNull(y) Then
            Equals = Null
        Else
            Equals = Empty
        End If
    End If
End Function

''' @param x As Variant(Of T)
''' @param y As Variant(Of T)
''' @return As Variant(Of Integer Or Null)
Public Function Compare(ByVal x As Variant, ByVal y As Variant) As Variant
    Dim xIsObj As Boolean: xIsObj = IsObject(x)
    Dim yIsObj As Boolean: yIsObj = IsObject(y)
    If xIsObj Xor yIsObj Then
        Err.Raise 13
    ElseIf xIsObj And yIsObj Then
        Compare = x.Compare(y)
    Else
        If TypeName(x) = TypeName(y) Then
            If x = y Then Compare = 0 Else _
            If x < y Then Compare = -1 Else _
            If x > y Then Compare = 1 Else _
            Compare = Null
        ElseIf IsNull(x) Or IsNull(y) Then
            Compare = Null
        Else
            Err.Raise 13
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
''' @return As Long
Public Function ArrLen(ByVal arr As Variant, Optional ByVal dimen As Integer = 1) As Long
    If Not IsArray(arr) Then Err.Raise 13
    ArrLen = UBound(arr, dimen) - LBound(arr, dimen) + 1
End Function

''' @param arr1 As Variant(Of Array(Of T))
''' @param arr2 As Variant(Of Array(Of T))
''' @return As Variant(Of Boolean Or Null)
Public Function ArrEquals(ByVal arr1 As Variant, ByVal arr2 As Variant) As Variant
    If Not (IsArray(arr1) And IsArray(arr2)) Then Err.Raise 13
    
    Dim alen1 As Long: alen1 = ArrLen(arr1)
    Dim alen2 As Long: alen2 = ArrLen(arr2)
    Dim cmpLen As Integer: cmpLen = Compare(alen1, alen2)
    
    Dim ix0 As Long: ix0 = LBound(arr1)
    Dim pad As Long: pad = LBound(arr2) - ix0
    Dim alenMin As Long: alenMin = IIf(cmpLen < 0, alen1, alen2)
    
    Dim i As Long, ret As Variant
    For i = ix0 To ix0 + alenMin - 1
        ret = Equals(arr1(i), arr2(pad + i))
        If ret Then Else GoTo Ending
    Next
    
    ret = Null
    Select Case cmpLen
    Case Is > 0
        For i = ix0 + alenMin To ix0 + alen1 - 1
            If IsNull(arr1(i)) Then GoTo Ending
        Next
    Case Is < 0
        For i = ix0 + alenMin To ix0 + alen2 - 1
            If IsNull(arr2(pad + i)) Then GoTo Ending
        Next
    End Select
    ret = Not CBool(cmpLen)
    
Ending:
    ArrEquals = ret
End Function

''' @param arr1 As Variant(Of Array(Of T))
''' @param arr2 As Variant(Of Array(Of T))
''' @return As Variant(Of Integer Or Null)
Public Function ArrCompare(ByVal arr1 As Variant, ByVal arr2 As Variant) As Variant
    If Not (IsArray(arr1) And IsArray(arr2)) Then Err.Raise 13
    
    Dim alen1 As Long: alen1 = ArrLen(arr1)
    Dim alen2 As Long: alen2 = ArrLen(arr2)
    Dim cmpLen As Integer: cmpLen = Compare(alen1, alen2)
    
    Dim ix0 As Long: ix0 = LBound(arr1)
    Dim pad As Long: pad = LBound(arr2) - ix0
    Dim alenMin As Long: alenMin = IIf(cmpLen < 0, alen1, alen2)
    
    Dim i As Long, ret As Variant
    For i = ix0 To ix0 + alenMin - 1
        ret = Compare(arr1(i), arr2(pad + i))
        If ret = 0 Then Else GoTo Ending
    Next
    
    ret = Null
    Select Case cmpLen
    Case Is > 0
        For i = ix0 + alenMin To ix0 + alen1 - 1
            If IsNull(arr1(i)) Then GoTo Ending
        Next
    Case Is < 0
        For i = ix0 + alenMin To ix0 + alen2 - 1
            If IsNull(arr2(pad + i)) Then GoTo Ending
        Next
    End Select
    ret = cmpLen
    
Ending:
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
        If Equals(arr(i), val) Then
            ArrIndexOf = i
            GoTo Escape
        End If
    Next
    
Escape:
End Function

''' @param arr As Variant(Of Array(Of T))
Public Sub ArrRev(ByRef arr As Variant)
    Dim ixL As Long: ixL = LBound(arr)
    Dim ixU As Long: ixU = UBound(arr)
    
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
            If Compare(arr(j - 1), arr(j)) * IIf(orderAsc, 1, -1) <= 0 Then Exit Do
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
            If Compare(arr(j - 1), arr(j)) * IIf(orderAsc, 1, -1) <= 0 Then Exit Do
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
    ArrMergeSw = Compare(arr1(i1), arr2(i2)) * IIf(orderAsc, 1, -1) < 1
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

''' @param fromVal As Variant(Of T)
''' @param toVal As Variant(Of T)
''' @param stepVal As Variant(Of T)
''' @return As Variant(Of Array(Of T))
Public Function ArrRange( _
    ByVal fromVal As Variant, ByVal toVal As Variant, Optional ByVal stepVal As Variant = 1 _
    ) As Variant
    
    If Not (IsNumeric(fromVal) And IsNumeric(toVal) And IsNumeric(stepVal)) Then Err.Raise 13
    
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

''' @param eobj As Enumerator(Of Object)
''' @return As Variant(Of Array(Of Object))
Public Function EnumToArr(ByVal eobj As Object) As Variant
    Dim i As Long: i = 0
    Dim alen As Long: alen = 32
    Dim arr As Variant: ReDim arr(alen - 1)
    
    Dim obj As Object
    For Each obj In eobj
        Set arr(IncrPst(i)) = obj
        If i >= alen Then alen = alen * 2: ReDim Preserve arr(alen - 1)
    Next
    
    If i > 0 Then
        ReDim Preserve arr(i - 1)
    Else
        arr = Array()
    End If
    EnumToArr = arr
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
    For i = 0 To UBound(aarr): Set aarr(i) = Init(New Tuple, arr(2 * i), arr(2 * i + 1)): Next
    
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
    For i = 0 To dlen: Set arr(i) = Init(New Tuple, ks(i), dict.Item(ks(i))): Next
    
Ending:
    DictToAssocArr = arr
End Function

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

''' @param num As Variant(Of Decimal)
''' @param digits As Integer
''' @param rndup As Integer
''' @return As Variant(Of Decimal)
Public Function ARound( _
    ByVal num As Variant, Optional ByVal digits As Integer = 0, Optional rndup As Integer = 5 _
    ) As Variant
    
    If IsNumeric(num) Then num = CDec(num) Else Err.Raise 13
    If Not (1 <= rndup And rndup <= 10) Then Err.Raise 5
    
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
   CharWidth = IIf((&H81 <= x And x <= &H9F) Or (&HE0 <= x And x <= &HFC), 2, 1)
End Function

''' @param s As String Is Char
''' @return As Long
Public Function StringWidth(ByVal s As String) As Long
    Dim w As Long: w = 0
    
    Dim i As Long
    For i = 1 To Len(s)
        w = w + CharWidth(Mid(s, i, 1))
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
        ixByte = ixByte + CharWidth(Mid(s, IncrPst(ixStr), 1))
    Wend
    LeftA = Left(s, ixStr - (ixByte - byteLen))
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
        ixByte = ixByte + CharWidth(Mid(s, IncrPst(ixStr), 1))
    Wend
    idxs.Add ixByte, ixStr
    
    For byteLen = byteLen To 0 Step -1
        If idxs.Exists(ixByte - byteLen) Then Exit For
    Next
    
    RightA = Right(s, ixStr - idxs.Item(ixByte - byteLen))
End Function

''' @param s As String
''' @param byteLen As Long
''' @return As Variant(Of Array(Of String))
Public Function SepA(ByVal s As String, ByVal byteLen As Long) As Variant
    Dim ixByte As Long: ixByte = 1
    Dim ixStr  As Long: ixStr = 1
    Dim strLen As Long: strLen = Len(s)
    
    While (ixByte < 1 + byteLen) And (ixStr <= strLen)
        ixByte = ixByte + CharWidth(Mid(s, IncrPst(ixStr), 1))
    Wend
    
    Dim n As Long: n = ixStr - (ixByte - byteLen)
    SepA = Array(Left(s, n), Mid(s, n + 1, strLen))
End Function
