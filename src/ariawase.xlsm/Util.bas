Attribute VB_Name = "Util"
Option Explicit
Option Private Module

''' @seealso WbemScripting.SWbemLocator http://msdn.microsoft.com/en-us/library/windows/desktop/aa393719.aspx

Public Enum HKeyEnum
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    'HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    'HKEY_DYN_DATA = &H80000006
End Enum

Private xxWmi As Object 'Is WbemScripting.SWbemLocator

'@return As Object Is WbemScripting.SWbemLocator
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

Public Function RegExpMatch( _
    ByVal expr As String, ByVal ptrnFind As String, _
    Optional ByVal regexpOption As String = "" _
    ) As String
    
    RegExpMatch = vbNullString
    
    Dim regex As Object: Set regex = CreateRegExp(ptrnFind, regexpOption)
    Dim ms As Object:    Set ms = regex.Execute(expr)
    If ms.Count > 0 Then RegExpMatch = ms.Item(0).Value
End Function

Public Function RegExpGlobalMatchs( _
    ByVal expr As String, ByVal ptrnFind As String, _
    Optional ByVal regexpOption As String = "g" _
    ) As Variant
    
    RegExpGlobalMatchs = Array()
    
    Dim regex As Object: Set regex = CreateRegExp(ptrnFind, regexpOption)
    Dim ms As Object:    Set ms = regex.Execute(expr)
    If ms.Count < 1 Then GoTo Escape
    
    Dim ret() As Variant
    ReDim ret(ms.Count - 1)
    Dim i As Long
    For i = 0 To UBound(ret): ret(i) = ms.Item(i).Value: Next
    RegExpGlobalMatchs = ret
    
Escape:
End Function

Public Function RegExpSubMatchs( _
    ByVal expr As String, ByVal ptrnFind As String, _
    Optional ByVal regexpOption As String = "" _
    ) As Variant
    
    RegExpSubMatchs = Array()
    
    Dim regex As Object: Set regex = CreateRegExp(ptrnFind, regexpOption)
    Dim ms As Object:    Set ms = regex.Execute(expr)
    If ms.Count < 1 Then GoTo Escape
    
    Dim sms As Object:   Set sms = ms(0).SubMatches
    If sms.Count < 1 Then GoTo Escape
    
    Dim ret() As Variant
    ReDim ret(sms.Count - 1)
    Dim i As Long
    For i = 0 To UBound(ret): ret(i) = sms.Item(i): Next
    RegExpSubMatchs = ret
    
Escape:
End Function

Public Function RegExpReplace( _
    ByVal expr As String, ByVal ptrnFind As String, ByVal ptrnReplace As String, _
    Optional ByVal regexpOption As String = "" _
    ) As String
    
    Dim regex As Object: Set regex = CreateRegExp(ptrnFind, regexpOption)
    RegExpReplace = regex.Replace(expr, ptrnReplace)
End Function

Public Function EvalVBS(ByVal vbsExpr As String) As Variant
    Dim sc As Object: Set sc = CreateObject("MSScriptControl.ScriptControl")
    sc.Language = "VBScript"
    EvalVBS = sc.Eval(vbsExpr)
End Function

Public Function EvalJS(ByVal jsExpr As String) As Variant
    Dim sc As Object: Set sc = CreateObject("MSScriptControl.ScriptControl")
    sc.Language = "JScript"
    EvalJS = sc.Eval(jsExpr)
End Function

''' @return As Object Is StdRegProv
Public Function CreateStdRegProv() As Object
    Dim wmiSrv As Object: Set wmiSrv = Wmi.ConnectServer(, "root\default")
    Set CreateStdRegProv = wmiSrv.Get("StdRegProv")
End Function
