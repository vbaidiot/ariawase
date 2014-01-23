Attribute VB_Name = "Ext"
''' +----                                                             --+ '''
''' |                          Ariawase 0.5.0                           | '''
''' |             Ariawase is free library for VBA cowboys.             | '''
''' |        The Project Page: https://github.com/igeta/Ariawase        | '''
''' +--                                                             ----+ '''
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
