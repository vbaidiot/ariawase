VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Func"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''+----                                                                   --+
'''|                             Ariawase 0.6.0 Beta                         |
'''|                Ariawase is free library for VBA cowboys.                |
'''|           The Project Page: https://github.com/igeta/Ariawase           |
'''+--                                                                   ----+
Option Explicit

#If VBA7 And Win64 Then
Private Declare PtrSafe _
Function DispCallFunc Lib "OleAut32.dll" ( _
    ByVal pvInstance As LongPtr, _
    ByVal oVft As LongPtr, _
    ByVal cc_ As Long, _
    ByVal vtReturn As Integer, _
    ByVal cActuals As Long, _
    ByRef prgvt As Integer, _
    ByRef prgpvarg As LongPtr, _
    ByRef pvargResult As Variant _
    ) As Long
#Else
Private Declare _
Function DispCallFunc Lib "OleAut32.dll" ( _
    ByVal pvInstance As Long, _
    ByVal oVft As Long, _
    ByVal cc_ As Long, _
    ByVal vtReturn As Integer, _
    ByVal cActuals As Long, _
    ByRef prgvt As Integer, _
    ByRef prgpvarg As Long, _
    ByRef pvargResult As Variant _
    ) As Long
#End If

Private Enum tagCALLCONV
    CC_FASTCALL = 0
    CC_CDECL = 1
    CC_MSCPASCAL = CC_CDECL + 1
    CC_PASCAL = CC_MSCPASCAL
    CC_MACPASCAL = CC_PASCAL + 1
    CC_STDCALL = CC_MACPASCAL + 1
    CC_FPFASTCALL = CC_STDCALL + 1
    CC_SYSCALL = CC_FPFASTCALL + 1
    CC_MPWCDECL = CC_SYSCALL + 1
    CC_MPWPASCAL = CC_MPWCDECL + 1
    CC_MAX = CC_MPWPASCAL
End Enum

Private Const S_OK = &H0

Private xxReturnType As VbVarType
#If VBA7 And Win64 Then
Private xxAddr As LongPtr
#Else
Private xxAddr As Long
#End If

Private xxParamTypeBinding As Boolean
Private xxUbParam As Long
Private xxParamTypes() As Integer

Private Sub EarlyBindParamTypes(ByVal paramTypes As Variant)
'    If Not IsArray(paramTypes) Then Err.Raise 5
    xxParamTypeBinding = False
    
    xxUbParam = UBound(paramTypes)
    
    If xxUbParam >= 0 Then
        Dim i As Long
        ReDim xxParamTypes(xxUbParam)
        For i = 0 To xxUbParam: xxParamTypes(i) = paramTypes(i): Next
        
        xxParamTypeBinding = True
    End If
End Sub

Private Sub LateBindParamTypes(ByVal params As Variant)
'    If Not IsArray(params) Then Err.Raise 5
    
    xxUbParam = UBound(params)
    
    If xxUbParam >= 0 Then
        Dim i As Long
        ReDim xxParamTypes(xxUbParam)
        For i = 0 To xxUbParam: xxParamTypes(i) = VarType(params(i)): Next
    Else
        ReDim xxParamTypes(0)
    End If
    
    xxParamTypeBinding = True
End Sub

#If VBA7 And Win64 Then
Public Sub Init( _
    ByVal retrunType As VbVarType, ByVal addr As LongPtr, ParamArray paramTypes() As Variant _
    )
#Else
Public Sub Init( _
    ByVal retrunType As VbVarType, ByVal addr As Long, ParamArray paramTypes() As Variant _
    )
#End If
    
    xxReturnType = retrunType
    xxAddr = addr
    EarlyBindParamTypes paramTypes
End Sub

Public Sub CallByPtr(ByRef returnValue As Variant, ByVal params As Variant)
'    If Not IsArray(params) Then Err.Raise 5
    
#If VBA7 And Win64 Then
    Dim lpParams() As LongPtr
#Else
    Dim lpParams() As Long
#End If
    If xxUbParam >= 0 Then
        Dim i As Long
        ReDim lpParams(xxUbParam)
        For i = 0 To xxUbParam: lpParams(i) = VarPtr(params(i)): Next
    Else
        ReDim lpParams(0)
    End If
    
    Dim stat As Long, ret As Variant
    stat = DispCallFunc( _
        0, xxAddr, tagCALLCONV.CC_STDCALL, xxReturnType, _
        xxUbParam + 1, xxParamTypes(0), lpParams(0), ret)
    If stat <> S_OK Then Err.Raise 5
    
    If IsObject(ret) Then Set returnValue = ret Else Let returnValue = ret
End Sub

Public Sub FastApply(ByRef returnValue As Variant, ParamArray params() As Variant)
    If Not xxParamTypeBinding Then
        LateBindParamTypes params
    End If
    CallByPtr returnValue, params
End Sub

Public Function Apply(ParamArray params() As Variant) As Variant
    If Not xxParamTypeBinding Then
        LateBindParamTypes params
    Else
        If xxUbParam <> UBound(params) Then Err.Raise 5
    End If
    CallByPtr Apply, params
End Function
