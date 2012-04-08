Attribute VB_Name = "Util"
Option Explicit
Option Private Module

''' @seealso WbemScripting.SWbemLocator http://msdn.microsoft.com/en-us/library/windows/desktop/aa393719.aspx
''' @seealso VBScript.RegExp http://msdn.microsoft.com/ja-jp/library/cc392403.aspx

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

''' @param expr As String
''' @param ptrnFind As String
''' @param iCase As Boolean
''' @return As Variant(Of Array(Of String))
Public Function RegExpMatch( _
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
    RegExpMatch = ret
End Function

''' @param expr As String
''' @param ptrnFind As String
''' @param iCase As Boolean
''' @return As Variant(Of Array(Of Array(Of String)))
Public Function RegExpGMatchs( _
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
    RegExpGMatchs = ret
End Function

''' @param expr As String
''' @param ptrnFind As String
''' @param ptrnReplace As String
''' @param regexpOption As String
''' @return As Variant(Of Array(Of String))
Public Function RegExpReplace( _
    ByVal expr As String, ByVal ptrnFind As String, ByVal ptrnReplace As String, _
    Optional ByVal regexpOption As String = "" _
    ) As String
    
    Dim regex As Object: Set regex = CreateRegExp(ptrnFind, regexpOption)
    RegExpReplace = regex.Replace(expr, ptrnReplace)
End Function

''' @usage
'''     Formats("{0:000} {{{1:yyyy/mm/dd}}} {2}", 1, Now, "Simple is best.") '001 {2012/04/08} Simple is best.
''' @param strTemplate As String
''' @param vals() As Variant
''' @return As String
Public Function Formats(ByVal strTemplate As String, ParamArray vals() As Variant) As String
    Dim ms As Variant: ms = RegExpGMatchs(strTemplate, "[^{]?({(\d)(:(.*?[^}]?))?})")
    Formats = Replace(Replace(strTemplate, "{{", "{"), "}}", "}")
    
    Dim m As Variant
    For Each m In ms
        Formats = Replace(Formats, m(1), Format(vals(m(2)), m(4)))
    Next
End Function

''' @param vbsExpr As String
''' @return As Variant
Public Function EvalVBS(ByVal vbsExpr As String) As Variant
    Dim sc As Object: Set sc = CreateObject("MSScriptControl.ScriptControl")
    sc.Language = "VBScript"
    EvalVBS = sc.Eval(vbsExpr)
End Function

''' @param jsExpr As String
''' @return As Variant
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
