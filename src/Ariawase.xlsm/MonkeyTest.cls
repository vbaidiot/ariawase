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
'''|                             Ariawase 0.9.0                              |
'''|                Ariawase is free library for VBA cowboys.                |
'''|          The Project Page: https://github.com/vbaidiot/Ariawase         |
'''+--                                                                   ----+
Option Explicit

' Assert.RunTestOf New MonkeyTest

Public Sub CorePrim_Test()
    Dim i As Integer, n As Long, v As Variant
    
    i = 42
    Assert.AreEqual 43, IncrPre(i)
    Assert.AreEqual 43, i
    Assert.IsInstanceOfTypeName "Integer", i
    n = 42
    Assert.AreEqual 43&, IncrPre(n)
    Assert.AreEqual 43&, n
    Assert.IsInstanceOfTypeName "Long", n
    v = CDec(42)
    Assert.AreEqual CDec(43), IncrPre(v)
    Assert.AreEqual CDec(43), v
    Assert.IsInstanceOfTypeName "Decimal", v
    
    i = 42
    Assert.AreEqual 42, IncrPst(i)
    Assert.AreEqual 43, i
    Assert.IsInstanceOfTypeName "Integer", i
    n = 42
    Assert.AreEqual 42&, IncrPst(n)
    Assert.AreEqual 43&, n
    Assert.IsInstanceOfTypeName "Long", n
    v = CDec(42)
    Assert.AreEqual CDec(42), IncrPst(v)
    Assert.AreEqual CDec(43), v
    Assert.IsInstanceOfTypeName "Decimal", v
    
    Assert.AreEqual 1&, BitFlag(False, True)
    Assert.AreEqual 8&, BitFlag(True, False, False, False)
    Assert.AreEqual 8&, BitFlag(1, 0, 0, 0)
    
    Assert.AreEqual True, IsInt(7#)
    Assert.AreEqual False, IsInt(7.01)
    Assert.AreEqual CDec(4), ARound(4.4)
    Assert.AreEqual CDec(5), ARound(4.5)
    Assert.AreEqual CDec(2.34), ARound(2.344, 2)
    Assert.AreEqual CDec(2.35), ARound(2.345, 2)
    Assert.AreEqual CDec(2), ARound(2.9, 0, 10)
    Assert.AreEqual CDec(3), ARound(2.1, 0, 1)
    
    Assert.AreEqual #5/1/2013#, BeginOfMonth(#5/15/2013#)
    Assert.AreEqual #5/31/2013#, EndOfMonth(#5/15/2013#)
    Assert.AreEqual #5/12/2013#, BeginOfWeek(#5/15/2013#)
    Assert.AreEqual #5/18/2013#, EndOfWeek(#5/15/2013#)
    
    '' the following test are for Japanese.
    Assert.AreEqual 13&, StringWidth("Testテストﾃｽﾄ")
    Assert.AreEqual 27&, StringWidth("全角2ﾊﾞｲﾄ換算による文字列長")
    Assert.AreEqual "全角2byte", LeftA("全角2byteで左から文字列を切り出す", 10)
    Assert.AreEqual "2byteで切り出す", RightA("右から文字列を全角2byteで切り出す", 16)
    Assert.AreEqualArr Array("LeftAと", "その余りﾀﾞﾖ"), SepA("LeftAとその余りﾀﾞﾖ", 8)
    
    Assert.AreEqual _
        "001 {2013/05/20} Simple is best.", _
        Formats("{0:000} {{{1:yyyy/mm/dd}}} {2}", 1, #5/20/2013#, "Simple is best.")
End Sub

Public Sub CoreAbst_Test()
    Assert.AreEqual Init(New Tuple, "A", 4), Init(New Tuple, "A", 4)
    Assert.AreEqual "A", Init(New Tuple, "A", 4).Item1
    Assert.AreEqual 4, Init(New Tuple, "A", 4).Item2
    
    Assert.AreEqual "Tuple`2 { Item1 = ""A"", Item2 = 4% }", ToStr(Init(New Tuple, "A", 4))
    Assert.AreEqual "Tuple`4 { Item1 = 1%, Item2 = 2&, Item3 = 3#, Item4 = 4@ }", ToStr(Init(New Tuple, 1, 2&, 3#, 4@))
    Assert.AreEqual "Variant()", ToStr(Array())
    
    Assert.AreEqual "42%", Dump(42)
    Assert.AreEqual "42&", Dump(42&)
    Assert.AreEqual "3.14#", Dump(3.14)
    Assert.AreEqual "#3/22/2014#", Dump(#3/22/2014#)
    Assert.AreEqual "#1/1/100#", Dump(#1/1/100#)
    Assert.AreEqual "#7:00:00 PM#", Dump(#7:00:00 PM#)
    Assert.AreEqual "#12/3/456 7:00:00 AM#", Dump(#12/3/456 7:00:00 AM#)
    Assert.AreEqual """Foo""""Bar""""Baz""", Dump("Foo""Bar""Baz")
    Assert.AreEqual """""", Dump("")
    Assert.AreEqual "vbNullString", Dump(vbNullString)
    Assert.AreEqual "Empty", Dump(Empty)
    Assert.AreEqual "Nothing", Dump(Nothing)
    Assert.AreEqual "Tuple`2 { Item1 = ""A"", Item2 = 4% }", Dump(Init(New Tuple, "A", 4))
    Assert.AreEqual "Array(1%, 2%, 3%, 4%, 5%, 6%, 7%, 8%, ...)", Dump(Array(1, 2, 3, 4, 5, 6, 7, 8, 9))
    Dim arr2D(0, 0) As Integer
    Assert.AreEqual "Integer(,)", Dump(arr2D)
    Dim arr3D(0, 0, 0) As Long
    Assert.AreEqual "Long(,,)", Dump(arr3D)
    
    Assert.AreEqual True, Eq(42, 42)
    Assert.AreEqual True, Eq(42, 42&)
    Assert.AreEqual False, Eq(42, "42")
    Assert.AreEqual True, Eq(Empty, 0)
    Assert.AreEqual True, Eq(Empty, Empty)
    Assert.AreEqual Null, Eq(Null, Empty)
    Assert.AreEqual Null, Eq(Null, Null)
    Assert.AreEqual True, Eq("", vbNullString)
    Assert.AreEqual False, Eq(Init(New Tuple, "A", 4), Init(New Tuple, "A", 4))
    Assert.AreEqual Empty, Eq(Init(New Tuple, "A", 4), "A4")
    
    Assert.AreEqual True, Equals(42, 42)
    Assert.AreEqual Empty, Equals(42, 42&)
    Assert.AreEqual Empty, Equals(42, "42")
    Assert.AreEqual Empty, Equals(Empty, 0)
    Assert.AreEqual True, Equals(Empty, Empty)
    Assert.AreEqual True, Equals("", vbNullString)
    Assert.AreEqual Null, Equals(Null, Empty)
    Assert.AreEqual Null, Equals(Null, Null)
    Assert.AreEqual True, Equals(Init(New Tuple, "A", 4), Init(New Tuple, "A", 4))
    Assert.AreEqual Empty, Equals(Init(New Tuple, "A", 4), "A4")
    
    Assert.AreEqual -1, Compare(3, 9)
    Assert.AreEqual 0, Compare(5, 5)
    Assert.AreEqual 1, Compare(9, 3)
    Assert.AreEqual Null, Compare(Null, Empty)
    Assert.AreEqual Null, Compare(Null, Null)
    Assert.IsErrFunc 13, Init(New Func, vbVariant, AddressOf Compare, vbVariant, vbVariant), Array(0, 0&)
    Assert.AreEqual -1, Compare(Init(New Tuple, 2, ".txt"), Init(New Tuple, 10, ".txt"))
    Assert.AreEqual 1, Compare(Init(New Tuple, "2", ".txt"), Init(New Tuple, "10", ".txt"))
    
    Assert.AreEqual 3, Min(3, 6, 5)
    Assert.IsInstanceOfTypeName "Integer", Min(3, 6, 5)
    Assert.AreEqual 6, Max(3, 6, 5)
    Assert.IsInstanceOfTypeName "Integer", Max(3, 6, 5)
End Sub

Public Sub CoreEnumr_Test()
    Dim arr2D(0, 0) As Integer
    Dim arr3D(0, 0, 0) As Long
    Assert.AreEqual 1, ArrRank(Array())
    Assert.AreEqual 2, ArrRank(arr2D)
    Assert.AreEqual 3, ArrRank(arr3D)
    
    Assert.AreEqual 3&, ArrLen(Array("V", "B", "A"))
    
    Assert.AreEqual True, ArrEquals(Array(0, 1, 2), Array(0, 1, 2))
    Assert.AreEqual False, ArrEquals(Array(2, 1, 0), Array(0, 1, 2))
    Assert.AreEqual Null, ArrEquals(Array(0, Null, 2), Array(0, Null, 2))
    Assert.AreEqual Empty, ArrEquals(Array(0, "1", 2), Array(0, 1, 2))
    Assert.AreEqual False, ArrEquals(Array(0, 1, 2, 3), Array(0, 1, 2))
    Assert.AreEqual False, ArrEquals(Array(0, 1, 2, Null), Array(0, 1, 2))
    Assert.AreEqual False, ArrEquals(Array(0, 1, 2, "3"), Array(0, 1, 2))
    Assert.AreEqual Null, ArrEquals(Array(Null, "1", 2, 3), Array(Null, 1, 2))
    Assert.AreEqual Empty, ArrEquals(Array("0", Null, 2, 3), Array(0, Null, 2))
    Assert.AreEqual True, ArrEquals(Array(), Array())
    Assert.AreEqual False, ArrEquals(Array(), Array(Empty))
    Assert.AreEqual True, ArrEquals(Array(0, 1&, "2"), Array(0, 1&, "2"))
    
    Assert.AreEqual True, ArrEquals(Array(0, Null, 2), Array(0, Null, 2), True)
    Assert.AreEqual False, ArrEquals(Array(Null, 1, 2), Array(0, 1, 2), True)
    Assert.AreEqual Empty, ArrEquals(Array(Null, "1", 2, 3), Array(Null, 1, 2), True)
    
    Assert.AreEqual 0, ArrCompare(Array(), Array())
    Assert.AreEqual -1, ArrCompare(Array(0, 1, 2), Array(2, 1, 0))
    Assert.AreEqual 0, ArrCompare(Array(0, 1, 2), Array(0, 1, 2))
    Assert.AreEqual 0, ArrCompare(Array(0, 1, 2), Array(0, 1, 2))
    Assert.AreEqual 1, ArrCompare(Array(2, 1, 0), Array(0, 1, 2))
    Assert.AreEqual Null, ArrCompare(Array(0, Null, 2), Array(0, Null, 2))
    Assert.IsErrFunc 13, _
        Init(New Func, vbVariant, AddressOf ArrCompare, vbVariant, vbVariant), _
        Array(Array(0, "1", 2), Array(0, 1, 2))
    Assert.AreEqual 1, ArrCompare(Array(0, 1, 2, 3), Array(0, 1, 2))
    Assert.AreEqual 1, ArrCompare(Array(0, 1, 2, Null), Array(0, 1, 2))
    Assert.AreEqual 1, ArrCompare(Array(0, 1, 2, "3"), Array(0, 1, 2))
    Assert.AreEqual Null, ArrCompare(Array(Null, "1", 2, 3), Array(Null, 1, 2))
    Assert.IsErrFunc 13, _
        Init(New Func, vbVariant, AddressOf ArrCompare, vbVariant, vbVariant), _
        Array(Array("0", Null, 2, 3), Array(0, Null, 2))
    Assert.AreEqual 0, ArrCompare(Array(), Array())
    Assert.AreEqual -1, ArrCompare(Array(), Array(Empty))
    Assert.AreEqual 0, ArrCompare(Array(0, 1&, "2"), Array(0, 1&, "2"))
    
    Assert.AreEqual 0, ArrCompare(Array(0, Null, 2), Array(0, Null, 2), True)
    Assert.AreEqual -1, ArrCompare(Array(Null, 1, 2), Array(0, 1, 2), True)
    Assert.IsErrFunc 13, _
        Init(New Func, vbVariant, AddressOf ArrCompare, vbVariant, vbVariant), _
        Array(Array(Null, "1", 2, 3), Array(Null, 1, 2), True)
    
    Assert.AreEqual 2&, ArrIndexOf(Array("V", "B", "A"), "A")
    Assert.AreEqual -1&, ArrIndexOf(Array("V", "B", "A"), "Z")
    Assert.AreEqual 1&, ArrIndexOf(Array("I", "I", "f"), "I", 1)
    Assert.AreEqual -1&, ArrIndexOf(Array("I", "I", "f"), "f", 0, 2)
    Assert.AreEqual 2&, ArrIndexOf(Array("I", "I", "f"), "f", 1, 9)
    Assert.AreEqual 3&, ArrIndexOf(Array(0, 1, 2, Null, 4, 5), Null)
    
    Dim arr As Variant
    
    arr = Array()
    ArrRev arr
    Assert.AreEqualArr Array(), arr
    arr = Array(2, 1)
    ArrRev arr
    Assert.AreEqualArr Array(1, 2), arr
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
    arr = Array(2, 1, 0, Null)
    ArrSort arr
    Assert.AreEqualArr Array(Null, 0, 1, 2), arr
    
    Assert.AreEqualArr Array(6, 5, 3), ArrUniq(Array(6, 5, 5, 3, 6))
    
    Assert.AreEqualArr Array(1, 2, 3, 4, 5), ArrConcat(Array(1, 2, 3), Array(4, 5))
    
    Assert.AreEqualArr Array(2, 3, 4), ArrSlice(Array(1, 2, 3, 4, 5), 1, 3)
    
    Assert.AreEqualArr Array(1, 2, 3, 4, 5), ArrFlatten(Array(Array(1, 2), Array(3), Array(4, 5)))
    
    Assert.AreEqualArr Array(1, 2, 3), ClctToArr(ArrToClct(Array(1, 2, 3)))
    
    Dim clct As Collection
    
    Set clct = ArrToClct(Array(0))
    Push clct, 3
    Assert.AreEqualArr Array(0, 3), ClctToArr(clct)
    Push clct, 4
    Assert.AreEqualArr Array(0, 3, 4), ClctToArr(clct)
    Assert.AreEqual 4, Pop(clct)
    Assert.AreEqualArr Array(0, 3), ClctToArr(clct)
    Shift clct, 2
    Assert.AreEqualArr Array(2, 0, 3), ClctToArr(clct)
    Shift clct, 1
    Assert.AreEqualArr Array(1, 2, 0, 3), ClctToArr(clct)
    Assert.AreEqual 1, Unshift(clct)
    Assert.AreEqualArr Array(2, 0, 3), ClctToArr(clct)
    
    Dim vArr2D As Variant
    vArr2D = JagArrToArr2D(Array(Array(1, 2, 3), Array(4, 5, 6)))
    Assert.AreEqual 2&, ArrLen(vArr2D, 1)
    Assert.AreEqual 3&, ArrLen(vArr2D, 2)
    Assert.AreEqual 1, vArr2D(0, 0)
    Assert.AreEqual 2, vArr2D(0, 1)
    Assert.AreEqual 3, vArr2D(0, 2)
    Assert.AreEqual 4, vArr2D(1, 0)
    Assert.AreEqual 5, vArr2D(1, 1)
    Assert.AreEqual 6, vArr2D(1, 2)
    
    Dim jagArr As Variant
    jagArr = Arr2DToJagArr(vArr2D)
    Assert.AreEqual 2&, ArrLen(jagArr)
    Assert.AreEqualArr Array(1, 2, 3), jagArr(0)
    Assert.AreEqualArr Array(4, 5, 6), jagArr(1)
    
    Dim dict As Object
    Set dict = CreateDictionary("One", 1, "Two", 2, "Three", 3)
    Assert.AreEqual 3&, dict.Count
    Assert.AreEqual 1, dict!One
    Assert.AreEqual 2, dict!Two
    Assert.AreEqual 3, dict!Three
End Sub

Public Sub CoreXLib_Test()
    Assert.AreEqual "C:\WINDOWS", UCase(Wsh.ExpandEnvironmentStrings("%windir%"))
    Assert.AreEqual 0&, Wmi.Security_.Privileges.Count
    
    Dim regex As Object
    Set regex = CreateRegExp(".*", "ig")
    Assert.AreEqual ".*", regex.Pattern
    Assert.AreEqual True, regex.Global
    Assert.AreEqual True, regex.IgnoreCase
    
    Assert.AreEqual "#VBA", ReMatch("beyond the limit of #VBA.", "#\w+")(0)
    Assert.AreEqualArr Array("2013/5/1", "2013", "5", "1"), ReMatch("2013/5/1", "(\d{4})/(\d{1,2})/(\d{1,2})")
    Assert.AreEqual 3&, ArrLen(ArrFlatten(ReMatcheGlobal("Practice! Practice! Practice!", "\w+!")))
    Assert.AreEqual _
        "<li>Word</li><li>Excel</li><li>Access</li>", _
        ReReplace("<p>Word</p><P>Excel</P><p>Access</p>", "<(/?)p>", "<$1li>", "ig")
    Assert.AreEqual _
        "WScript.Echo(""Hello!"")", _
        ReTrim(vbTab & "WScript.Echo(""Hello!"");" & vbCrLf, "[\s;]")
    
    Assert.AreEqual "7", EvalVBS("""3"" + 4")
    Assert.AreEqual "34", EvalJS("""3"" + 4")
    
    Dim s As String
    CreateStdRegProv().GetStringValue HKEY_CLASSES_ROOT, ".txt", "Content Type", s
    Assert.AreEqual "text/plain", s
End Sub

Public Sub Ext_Test()
    Dim aarr As Variant, dict As Object
    
    aarr = CreateAssocArray("One", 1, "Two", 2, "Three", 3)
    Assert.AreEqual 3&, ArrLen(aarr)
    Assert.AreEqual Init(New Tuple, "One", 1), aarr(0)
    Assert.AreEqual Init(New Tuple, "Two", 2), aarr(1)
    Assert.AreEqual Init(New Tuple, "Three", 3), aarr(2)
    
    Set dict = AssocArrToDict(aarr)
    Assert.AreEqual 3&, dict.Count
    Assert.AreEqual 1, dict!One
    Assert.AreEqual 2, dict!Two
    Assert.AreEqual 3, dict!Three
    
    aarr = DictToAssocArr(dict)
    Assert.AreEqual 3&, ArrLen(aarr)
    Assert.AreEqual Init(New Tuple, "One", 1), aarr(0)
    Assert.AreEqual Init(New Tuple, "Two", 2), aarr(1)
    Assert.AreEqual Init(New Tuple, "Three", 3), aarr(2)
    
    Assert.IsInstanceOfTypeName _
        "Decimal", _
        CDec(EnumeratorToArr(Wmi.ConnectServer().ExecQuery("SELECT * FROM Win32_ComputerSystem"))(0).TotalPhysicalMemory)
    
    Assert.AreEqualArr Array(1, 2, 3, 4, 5, 6, 7, 8, 9), ArrRange(1, 9)
    
    Assert.AreEqualArr _
        Array("1", "2", "3"), _
        ArrMap(Init(New Func, vbString, AddressOf ToStr, vbVariant), Array(1, 2, 3))
    
'    Assert.AreEqualArr _
'        Array(Init(New Tuple, "A", 1), Init(New Tuple, "B", 2)), _
'        ArrZip(Init(New Func, vbObject, AddressOf Tuple2Of, vbVariant, vbVariant), Array("A", "B"), Array(1, 2))
    
'    Assert.AreEqualArr _
'        Array(1, 3, 5, 7, 9), _
'        ArrFilter(Init(New Func, vbBoolean, AddressOf IsOdd, vbInteger), ArrRange(1, 9))
    
'    Dim lkup As Variant
'    lkup = ArrGroupBy( _
'        Init(New Func, vbVariant, AddressOf GetItem1, vbObject), _
'        Array( _
'            Init(New Tuple, "A", 1), Init(New Tuple, "B", 4), Init(New Tuple, "C", 3), _
'            Init(New Tuple, "A", 2), Init(New Tuple, "B", 3), Init(New Tuple, "C", 1), _
'            Init(New Tuple, "A", 3), Init(New Tuple, "B", 2), Init(New Tuple, "C", 2), _
'            Init(New Tuple, "A", 4), Init(New Tuple, "B", 1)))
'    Assert.AreEqual 3&, ArrLen(lkup)
'    Assert.AreEqual "A", lkup(0).Item1
'    Assert.AreEqual "B", lkup(1).Item1
'    Assert.AreEqual "C", lkup(2).Item1
'    Assert.AreEqualArr _
'        Array(Init(New Tuple, "A", 1), Init(New Tuple, "A", 2), Init(New Tuple, "A", 3), Init(New Tuple, "A", 4)), _
'        lkup(0).Item2
'    Assert.AreEqualArr _
'        Array(Init(New Tuple, "B", 4), Init(New Tuple, "B", 3), Init(New Tuple, "B", 2), Init(New Tuple, "B", 1)), _
'        lkup(1).Item2
'    Assert.AreEqualArr _
'        Array(Init(New Tuple, "C", 3), Init(New Tuple, "C", 1), Init(New Tuple, "C", 2)), _
'        lkup(2).Item2
    
'    Assert.AreEqual _
'        5050, _
'        ArrFold(Init(New Func, vbInteger, AddressOf Add, vbInteger, vbInteger), ArrRange(1, 100), 0)
    
'    Assert.AreEqualArr _
'        Array(0, 1, 3, 6, 10, 15, 21, 28, 36, 45, 55), _
'        ArrScan(Init(New Func, vbInteger, AddressOf Add, vbInteger, vbInteger), ArrRange(1, 10), 0)
    
'    Assert.AreEqualArr _
'        Array(1&, 2&, 3&, 5&, 8&, 13&, 21&, 34&, 55&, 89&), _
'        ArrUnfold(Init(New Func, vbVariant, AddressOf FibFun, vbObject), Init(New Tuple, 1, 1))
End Sub

#If False Then
''' Note: copy the following codes paste onto a standard module

Public Function Tuple2Of(ByVal itm1 As Variant, ByVal itm2 As Variant) As Tuple
    Set Tuple2Of = Init(New Tuple, itm1, itm2)
End Function

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
