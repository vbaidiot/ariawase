VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Func"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

#If VBA7 And Win64 Then
Private xxAddr As LongPtr
#Else
Private xxAddr As Long
#End If
Private xxReturnType As VbVarType
Private xxParamTypes() As Variant

#If VBA7 And Win64 Then
Public Sub Init(ByVal addr As LongPtr, ByVal retrunType As VbVarType, ParamArray paramTypes() As Variant)
#Else
Public Sub Init(ByVal addr As Long, ByVal retrunType As VbVarType, ParamArray paramTypes() As Variant)
#End If
    xxAddr = addr
    xxReturnType = retrunType
    xxParamTypes = paramTypes
End Sub

Public Sub Apply(ByRef ret As Variant, ParamArray params() As Variant)
    Dim ubParam As Long: ubParam = UBound(params)
#If VBA7 And Win64 Then
    Dim ps() As LongPtr: ReDim ps(ubParam)
#Else
    Dim ps() As Long: ReDim ps(ubParam)
#End If
    
    Dim i As Integer
    Select Case UBound(xxParamTypes)
    Case -1
        ReDim xxParamTypes(ubParam)
        For i = 0 To ubParam: xxParamTypes(i) = VarType(params(i)): Next
    Case Is <> ubParam
        Err.Raise 5
    Case Else
        ' DO NOTHING
    End Select
    
    For i = 0 To UBound(ps): ps(i) = VarPtr(params(i)): Next
    
    Dim stat As Long
    stat = DispCallFunc( _
        0, xxAddr, tagCALLCONV.CC_STDCALL, xxReturnType, _
        ubParam + 1, xxParamTypes(0), ps(0), ret)
    If stat <> S_OK Then Err.Raise 5
End Sub
