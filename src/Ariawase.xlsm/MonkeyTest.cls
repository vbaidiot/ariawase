VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MonkeyTest"
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

' Assert.RunTestOf New MonkeyTest

Public Sub Core_Test()
    Dim i As Integer, n As Long, v As Variant
    
    i = 42
    Assert.AreEqualVal 43, IncrPre(i)
    Assert.AreEqualVal 43, i
    Assert.IsInstanceOfTypeName "Integer", i
    n = 42
    Assert.AreEqualVal 43&, IncrPre(n)
    Assert.AreEqualVal 43&, n
    Assert.IsInstanceOfTypeName "Long", n
    v = CDec(42)
    Assert.AreEqualVal CDec(43), IncrPre(v)
    Assert.AreEqualVal CDec(43), v
    Assert.IsInstanceOfTypeName "Decimal", v
    
    i = 42
    Assert.AreEqualVal 42, IncrPst(i)
    Assert.AreEqualVal 43, i
    Assert.IsInstanceOfTypeName "Integer", i
    n = 42
    Assert.AreEqualVal 42&, IncrPst(n)
    Assert.AreEqualVal 43&, n
    Assert.IsInstanceOfTypeName "Long", n
    v = CDec(42)
    Assert.AreEqualVal CDec(42), IncrPst(v)
    Assert.AreEqualVal CDec(43), v
    Assert.IsInstanceOfTypeName "Decimal", v
    
    Assert.AreEqualVal 1&, BitFlag(False, True)
    Assert.AreEqualVal 8&, BitFlag(True, False, False, False)
    Assert.AreEqualVal 8&, BitFlag(1, 0, 0, 0)
    
    Assert.AreEqualVal CDec(4), ARound(4.4)
    Assert.AreEqualVal CDec(5), ARound(4.5)
    Assert.AreEqualVal CDec(2.34), ARound(2.344, 2)
    Assert.AreEqualVal CDec(2.35), ARound(2.345, 2)
    Assert.AreEqualVal CDec(2), ARound(2.9, 0, 10)
    Assert.AreEqualVal CDec(3), ARound(2.1, 0, 1)
    
    Assert.AreEqualVal #5/1/2013#, BeginOfMonth(#5/15/2013#)
    Assert.AreEqualVal #5/31/2013#, EndOfMonth(#5/15/2013#)
    Assert.AreEqualVal #5/12/2013#, BeginOfWeek(#5/15/2013#)
    Assert.AreEqualVal #5/18/2013#, EndOfWeek(#5/15/2013#)
    
    Assert.AreEqualVal Core.Init(New Tuple, "A", 4), Core.Init(New Tuple, "A", 4)
    Assert.AreEqualVal "A", Core.Init(New Tuple, "A", 4).Item1
    Assert.AreEqualVal 4, Core.Init(New Tuple, "A", 4).Item2
    Assert.AreEqualVal "Tuple`2 { Item1 = ""A"", Item2 = 4% }", ToStr(Core.Init(New Tuple, "A", 4))
    Assert.AreEqualVal "Tuple`4 { Item1 = 1%, Item2 = 2&, Item3 = 3#, Item4 = 4@ }", ToStr(Core.Init(New Tuple, 1, 2&, 3#, 4@))
    
    Assert.AreEqualVal True, Eq(42, 42)
    Assert.AreEqualVal True, Eq(42, 42&)
    Assert.AreEqualVal False, Eq(42, "42")
    Assert.AreEqualVal True, Eq(Empty, 0)
    Assert.AreEqualVal True, Eq(Empty, Empty)
    Assert.IsNullVal Eq(Null, Empty)
    Assert.IsNullVal Eq(Null, Null)
    Assert.AreEqualVal True, Eq("", vbNullString)
    Assert.AreEqualVal False, Eq(Core.Init(New Tuple, "A", 4), Core.Init(New Tuple, "A", 4))
    Assert.AreEqualVal Empty, Eq(Core.Init(New Tuple, "A", 4), "A4")
    
    Assert.AreEqualVal True, Equals(42, 42)
    Assert.AreEqualVal Empty, Equals(42, 42&)
    Assert.AreEqualVal Empty, Equals(42, "42")
    Assert.AreEqualVal Empty, Equals(Empty, 0)
    Assert.AreEqualVal True, Equals(Empty, Empty)
    Assert.AreEqualVal True, Equals("", vbNullString)
    Assert.IsNullVal Equals(Null, Empty)
    Assert.IsNullVal Equals(Null, Null)
    Assert.AreEqualVal True, Equals(Core.Init(New Tuple, "A", 4), Core.Init(New Tuple, "A", 4))
    Assert.AreEqualVal Empty, Equals(Core.Init(New Tuple, "A", 4), "A4")
    
    Assert.AreEqualVal -1, Compare(3, 9)
    Assert.AreEqualVal 0, Compare(5, 5)
    Assert.AreEqualVal 1, Compare(9, 3)
    Assert.IsNullVal Compare(Null, Empty)
    Assert.IsNullVal Compare(Null, Null)
    Assert.IsErrFunc 13, Core.Init(New Func, vbVariant, AddressOf Compare, vbVariant, vbVariant), Array(0, 0&)
    Assert.AreEqualVal -1, Compare(Core.Init(New Tuple, 2, ".txt"), Core.Init(New Tuple, 10, ".txt"))
    Assert.AreEqualVal 1, Compare(Core.Init(New Tuple, "2", ".txt"), Core.Init(New Tuple, "10", ".txt"))
    
    Assert.AreEqualVal 3, Min(3, 6, 5)
    Assert.IsInstanceOfTypeName "Integer", Min(3, 6, 5)
    Assert.AreEqualVal 6, Max(3, 6, 5)
    Assert.IsInstanceOfTypeName "Integer", Max(3, 6, 5)
    
    Assert.AreEqualVal 3&, ArrLen(Array("V", "B", "A"))
    
    Assert.AreEqualVal True, ArrEquals(Array(0, 1, 2), Array(0, 1, 2))
    Assert.AreEqualVal False, ArrEquals(Array(0, 1, 2), Array(2, 1, 0))
    Assert.IsNullVal ArrEquals(Array(Null), Array(1, 2, 3))
    Assert.IsNullVal ArrEquals(Array(1, 2, 3, Null), Array(1, 2, 3))
    
    Assert.AreEqualVal 0, ArrCompare(Array(0, 1, 2), Array(0, 1, 2))
    Assert.AreEqualVal -1, ArrCompare(Array(0, 1, 2), Array(0, 1, 3))
    Assert.AreEqualVal -1, ArrCompare(Array(0, 1, 2), Array(0, 1, 2, 3))
    Assert.AreEqualVal 1, ArrCompare(Array(2, 1), Array(0, 1, 2))
    Assert.IsNullVal ArrCompare(Array(Null), Array(1, 2, 3))
    Assert.IsNullVal ArrCompare(Array(1, 2, 3, Null), Array(1, 2, 3))
    
    Assert.AreEqualVal 2&, ArrIndexOf(Array("V", "B", "A"), "A")
    Assert.AreEqualVal -1&, ArrIndexOf(Array("V", "B", "A"), "Z")
    Assert.AreEqualVal 1&, ArrIndexOf(Array("I", "I", "f"), "I", 1)
    Assert.AreEqualVal -1&, ArrIndexOf(Array("I", "I", "f"), "f", 0, 2)
    Assert.AreEqualVal 2&, ArrIndexOf(Array("I", "I", "f"), "f", 1, 9)
    
    Dim arr As Variant
    
    arr = Array(1, 2, 3, 4, 5, 6, 7, 8, 9)
    ArrRev arr
    Assert.AreEqualArr Array(9, 8, 7, 6, 5, 4, 3, 2, 1), arr
    
    arr = Array("S", "O", "R", "T")
    ArrSort arr
    Assert.AreEqualArr Array("O", "R", "S", "T"), arr
    arr = Array(2, 4, 3, 5, 6, 7, 0, 1, 9, 8)
    ArrSort arr, True
    Assert.AreEqualArr Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9), arr
    arr = Array(2, 4, 3, 5, 6, 7, 0, 1, 9, 8)
    ArrSort arr, False
    Assert.AreEqualArr Array(9, 8, 7, 6, 5, 4, 3, 2, 1, 0), arr
    
    Assert.AreEqualArr Array(6, 5, 3), ArrUniq(Array(6, 5, 5, 3, 6))
    
    Assert.AreEqualArr Array(1, 2, 3, 4, 5), ArrConcat(Array(1, 2, 3), Array(4, 5))
    
    Assert.AreEqualArr Array(2, 3, 4), ArrSlice(Array(1, 2, 3, 4, 5), 1, 3)
    
    Assert.AreEqualArr Array(1, 2, 3, 4, 5), ArrFlatten(Array(Array(1, 2), Array(3), Array(4, 5)))
    
    Assert.AreEqualArr _
        Array("1", "2", "3"), _
        ArrMap(Core.Init(New Func, vbString, AddressOf ToStr, vbVariant), Array(1, 2, 3))
    
    Assert.AreEqualArr Array(1, 2, 3), ClctToArr(ArrToClct(Array(1, 2, 3)))
    
    Dim clct As Collection
    
    Set clct = New Collection
    Push clct, 3
    Assert.AreEqualArr Array(3), ClctToArr(clct)
    Push clct, 4
    Assert.AreEqualArr Array(3, 4), ClctToArr(clct)
    Assert.AreEqualVal 4, Pop(clct)
    Assert.AreEqualArr Array(3), ClctToArr(clct)
    Shift clct, 2
    Assert.AreEqualArr Array(2, 3), ClctToArr(clct)
    Shift clct, 1
    Assert.AreEqualArr Array(1, 2, 3), ClctToArr(clct)
    Assert.AreEqualVal 1, Unshift(clct)
    Assert.AreEqualArr Array(2, 3), ClctToArr(clct)
    
'    JagArrToArr2D
'    Arr2DToJagArr
'    CreateDictionary
    
    Assert.AreEqualVal "C:\WINDOWS", UCase(Wsh.ExpandEnvironmentStrings("%windir%"))
    
'    CreateRegExp
    
    Assert.AreEqualVal "#VBA", ReMatch("beyond the limit of #VBA.", "#\w+")(0)
    Assert.AreEqualArr Array("2013/5/1", "2013", "5", "1"), ReMatch("2013/5/1", "(\d{4})/(\d{1,2})/(\d{1,2})")
    Assert.AreEqualVal 3&, ArrLen(ArrFlatten(ReMatcheGlobal("Practice! Practice! Practice!", "\w+!")))
    Assert.AreEqualVal _
        "<li>Word</li><li>Excel</li><li>Access</li>", _
        ReReplace("<p>Word</p><P>Excel</P><p>Access</p>", "<(/?)p>", "<$1li>", "ig")
    
    Assert.AreEqualVal _
        "001 {2013/05/20} Simple is best.", _
        Formats("{0:000} {{{1:yyyy/mm/dd}}} {2}", 1, #5/20/2013#, "Simple is best.")
    
    Assert.AreEqualVal "7", EvalVBS("""3"" + 4")
    Assert.AreEqualVal "34", EvalJS("""3"" + 4")
    
    Dim s As String
    CreateStdRegProv().GetStringValue HKEY_CLASSES_ROOT, ".txt", "Content Type", s
    Assert.AreEqualVal "text/plain", s
End Sub

Public Sub Util_Test()
'    CreateAssocArray
'    AssocArrToDict
'    DictToAssocArr
    
    Assert.IsInstanceOfTypeName _
        "Decimal", _
        CDec(EnumeratorToArr(Wmi.ConnectServer().ExecQuery("SELECT * FROM Win32_ComputerSystem"))(0).TotalPhysicalMemory)
    
    Assert.AreEqualArr Array(1, 2, 3, 4, 5, 6, 7, 8, 9), ArrRange(1, 9)
    
'    Assert.AreEqualArr _
'        Array(1, 3, 5, 7, 9), _
'        ArrFilter(Init(New Func, vbBoolean, AddressOf IsOdd, vbInteger), ArrRange(1, 9))
    
'    '' FIXME: This test does not work.
'    Assert.AreEqualArr _
'        Array( _
'            Init(New Tuple, "A", Array( _
'                Init(New Tuple, "A", 1), Init(New Tuple, "A", 2), Init(New Tuple, "A", 3), Init(New Tuple, "A", 4) _
'                )), _
'            Init(New Tuple, "B", Array( _
'                Init(New Tuple, "B", 4), Init(New Tuple, "B", 3), Init(New Tuple, "B", 2), Init(New Tuple, "B", 1) _
'                )), _
'            Init(New Tuple, "C", Array( _
'                Init(New Tuple, "C", 3), Init(New Tuple, "C", 1), Init(New Tuple, "C", 2) _
'                ))), _
'        ArrGroupBy( _
'            Init(New Func, vbVariant, AddressOf GetItem1, vbObject), _
'            Array( _
'                Init(New Tuple, "A", 1), Init(New Tuple, "B", 4), Init(New Tuple, "C", 3), _
'                Init(New Tuple, "A", 2), Init(New Tuple, "B", 3), Init(New Tuple, "C", 1), _
'                Init(New Tuple, "A", 3), Init(New Tuple, "B", 2), Init(New Tuple, "C", 2), _
'                Init(New Tuple, "A", 4), Init(New Tuple, "B", 1)))
    
'    Assert.AreEqualVal _
'        5050, _
'        ArrFold(Init(New Func, vbInteger, AddressOf Add, vbInteger, vbInteger), ArrRange(1, 100), 0)
    
'    Assert.AreEqualArr _
'        Array(1&, 2&, 3&, 5&, 8&, 13&, 21&, 34&, 55&, 89&), _
'        ArrUnfold(Init(New Func, vbVariant, AddressOf FibFun, vbObject), Init(New Tuple, 1, 1))
End Sub

#If False Then
''' Note: copy the following codes paste onto a standard module
Public Function IsOdd(ByVal n As Integer) As Boolean
    IsOdd = n Mod 2 = 1
End Function

Public Function Add(ByVal x As Integer, ByVal y As Integer) As Integer
    Add = x + y
End Function

Public Function GetItem1(ByVal tpl As Tuple) As Variant
    GetItem1 = tpl.Item1
End Function

Public Function FibFun(ByVal tpl As Tuple) As Variant
    Dim lmt As Long: lmt = 100
    
    Dim v1 As Long: v1 = tpl.Item2
    Dim v2 As Long: v2 = tpl.Item1 + v1
    If v1 <= lmt Then
        Set FibFun = Init(New Tuple, v1, Init(New Tuple, v1, v2))
    Else
        Let FibFun = Missing
    End If
End Function
#End If
