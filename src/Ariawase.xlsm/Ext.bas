Attribute VB_Name = "Ext"
'''+----                                                                   --+
'''|                             Ariawase 0.6.0 Beta                         |
'''|                Ariawase is free library for VBA cowboys.                |
'''|           The Project Page: https://github.com/igeta/Ariawase           |
'''+--                                                                   ----+
Option Explicit
Option Private Module

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

Public Function AssocArrToDict(ByVal aarr As Variant) As Object
    If Not IsArray(aarr) Then Err.Raise 13
    Set AssocArrToDict = CreateDictionary()
    Dim v As Variant '(Of Tuple`2)
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

''' @param enumr As Enumerator(Of T)
''' @return As Variant(Of Array(Of T))
Public Function EnumeratorToArr(ByVal enumr As Object) As Variant
    Dim arrx As ArrayEx: Set arrx = New ArrayEx
    
    Dim x As Object
    For Each x In enumr: Exit For: Next
    If IsObject(x) Then
        For Each x In enumr: arrx.AddObj x: Next
    Else
        For Each x In enumr: arrx.AddVal x: Next
    End If
    
    EnumeratorToArr = arrx.ToArray()
End Function

''' @param fromVal As Variant(Of T)
''' @param toVal As Variant(Of T)
''' @param stepVal As Variant(Of T)
''' @return As Variant(Of Array(Of T))
Public Function ArrRange( _
    ByVal fromVal As Variant, ByVal toVal As Variant, Optional ByVal stepVal As Variant = 1 _
    ) As Variant
    
    If Not (IsNumeric(fromVal) And IsNumeric(toVal) And IsNumeric(stepVal)) Then Err.Raise 13
    
    Dim arrx As ArrayEx: Set arrx = New ArrayEx
    
    Select Case stepVal
    Case Is > 0
        While fromVal <= toVal
            arrx.AddVal IncrPst(fromVal, stepVal)
        Wend
    Case Is < 0
        While fromVal >= toVal
            arrx.AddVal IncrPst(fromVal, stepVal)
        Wend
    Case Else
        Err.Raise 5
    End Select
    
    ArrRange = arrx.ToArray()
End Function

''' @param fun As Func(Of T, U)
''' @param arr As Variant(Of Array(Of T))
''' @return As Variant(Of Array(Of U))
Public Function ArrMap(ByVal fun As Func, ByVal arr As Variant) As Variant
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
    For i = lb To ub: fun.FastApply ret(i), arr(i): Next
    
Ending:
    ArrMap = ret
End Function

''' @param fun As Func(Of T, Boolean)
''' @param arr As Variant(Of Array(Of T))
''' @return As Variant(Of Array(Of T))
Public Function ArrFilter(ByVal fun As Func, ByVal arr As Variant) As Variant
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
            fun.FastApply flg, arr(ixArr)
            If flg Then Set ret(IncrPst(ixRet)) = arr(ixArr)
        Next
    Else
        For ixArr = lb To ub
            fun.FastApply flg, arr(ixArr)
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

''' @param fun As Func(Of T, K)
''' @param arr As Variant(Of Array(Of T))
''' @return As Variant(Of Array(Of Tuple`2(Of K, T)))
Public Function ArrGroupBy(ByVal fun As Func, ByVal arr As Variant) As Variant
    If Not IsArray(arr) Then Err.Raise 13
    Dim lb As Long: lb = LBound(arr)
    Dim ub As Long: ub = UBound(arr)
    Dim ixRet As Long: ixRet = -1
    Dim ret As Variant
    If ub - lb < 0 Then
        ret = Array()
        GoTo Ending
    End If
    
    ReDim ret(ub - lb)
    
    Dim k As Variant, i As Long, j As Long
    If IsObject(arr(lb)) Then
        For i = lb To ub
            fun.FastApply k, arr(i)
            For j = ixRet To 0 Step -1
                If Equals(k, ret(j)(0)) Then Exit For
            Next
            If j < 0 Then
                j = IncrPre(ixRet)
                ret(j) = Array(k, New ArrayEx)
            End If
            ret(j)(1).AddObj arr(i)
        Next
    Else
        For i = lb To ub
            fun.FastApply k, arr(i)
            For j = ixRet To 0 Step -1
                If Equals(k, ret(j)(0)) Then Exit For
            Next
            If j < 0 Then
                j = IncrPre(ixRet)
                ret(j) = Array(k, New ArrayEx)
            End If
            ret(j)(1).AddVal arr(i)
        Next
    End If
    
    ReDim Preserve ret(ixRet)
    
    For i = 0 To ixRet
        Set ret(i) = Init(New Tuple, ret(i)(0), ret(i)(1).ToArray())
    Next
    
Ending:
    ArrGroupBy = ret
End Function

''' @param fun As Func(Of U, T, U)
''' @param arr As Variant(Of Array(Of T))
''' @param seedVal As Variant(Of U)
''' @return As Variant(Of U)
Public Function ArrFold( _
    ByVal fun As Func, ByVal arr As Variant, Optional ByVal seedVal As Variant _
    ) As Variant
    
    If Not IsArray(arr) Then Err.Raise 13
    
    Dim stat As Variant
    Dim i As Long: i = LBound(arr)
    If IsMissing(seedVal) Then
        stat = arr(IncrPst(i))
    Else
        stat = seedVal
    End If
    
    For i = i To UBound(arr): fun.FastApply stat, stat, arr(i): Next
    
    If IsObject(stat) Then
        Set ArrFold = stat
    Else
        Let ArrFold = stat
    End If
End Function

''' @param fun As Func
''' @param seedVal As Variant(Of T)
''' @return As Variant(Of Array(Of U))
Public Function ArrUnfold(ByVal fun As Func, ByVal seedVal As Variant) As Variant
    Dim arrx As ArrayEx: Set arrx = New ArrayEx
    
    Dim stat As Variant '(Of Tuple`2 Or Missing)
    fun.FastApply stat, seedVal
    If IsMissing(stat) Then GoTo Ending
    
    If IsObject(stat.Item1) Then
        arrx.AddObj stat.Item1
        
        fun.FastApply stat, stat.Item2
        While Not IsMissing(stat)
            arrx.AddObj stat.Item1
            fun.FastApply stat, stat.Item2
        Wend
    Else
        arrx.AddVal stat.Item1
        
        fun.FastApply stat, stat.Item2
        While Not IsMissing(stat)
            arrx.AddVal stat.Item1
            fun.FastApply stat, stat.Item2
        Wend
    End If
    
Ending:
    ArrUnfold = arrx.ToArray()
End Function
