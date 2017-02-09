VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IsJagArrTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Assert.RunTestOf New IsJagArrTest

Public Sub IsJagArr_Test()

    'A: Jagged Array
    Dim arr0 As Variant: arr0 = Array(Array(1))
    Dim arr1 As Variant: arr1 = Array(Array(1, 2, 3), Array(4, 5, 6))
    Dim arr2 As Variant: arr2 = Array(Array(1, 2, 3), Array(4, 5, 6), Array(7, 8, 9))
    Dim arr3 As Variant: arr3 = Array(Array(1), Array(2, 3, 4), Array(5, 6))
    Dim arr4 As Variant: arr4 = Array(Array(), Array(1, 2, 3), Array(4, 5))
    
    Assert.AreEqual True, IsJagArr(arr0), "arr0"
    Assert.AreEqual True, IsJagArr(arr1), "arr1"
    Assert.AreEqual True, IsJagArr(arr2), "arr2"
    Assert.AreEqual True, IsJagArr(arr3), "arr3"
    Assert.AreEqual True, IsJagArr(arr4), "arr4"


    'B: Array
    Dim arrB0 As Variant: arrB0 = Array(Array(), Array())
    Dim arrB1 As Variant: arrB1 = Array(ArrToClct(Array()), ArrToClct(Array(0)), ArrToClct(Array(0, 0)))
    Dim arrB2 As Variant: arrB2 = JagArrToArr2D(arr0)
    Dim arrB3 As Variant: arrB3 = JagArrToArr2D(arr1)
    
    Assert.AreEqual False, IsJagArr(arrB0), "arrB0"
    Assert.AreEqual False, IsJagArr(arrB1), "arrB1"
    Assert.AreEqual False, IsJagArr(arrB2), "arrB2"
    Assert.AreEqual False, IsJagArr(arrB3), "arrB3"
    
    
    'C: Dimention Error 9
    Dim arrC0() As Variant
    Assert.AreEqual False, IsJagArr(arrC0), "C0"
    
    
    'D: Dimention < 0
    Dim arrD0(-1 To -1) As Variant: arrD0(-1) = Array(0)
    Dim arrD1(-3 To -1) As Variant
        arrD1(-3) = Array(0)
        arrD1(-2) = Array(1)
        arrD1(-1) = Array(2)
    Dim arrD2(-3 To -1) As Variant
        arrD2(-3) = Array()
        arrD2(-2) = Array(0)
        arrD2(-1) = Array(ArrToClct(Array()))
    
    Assert.AreEqual True, IsJagArr(arrD0), "D0"  'arr(-1)(0)
    Assert.AreEqual True, IsJagArr(arrD1), "D1"  'arr(-3)(0), arr(-2)(0), arr(-1)(0)
    Assert.AreEqual True, IsJagArr(arrD2), "D2"  'arr(-3)(0), arr(-2)(0), arr(-1)(0)

    Dim arrD3(-1 To -1) As Variant 'value is `Empty`
    Dim arrD4(-1 To -1) As Variant: arrD4(-1) = 0
    Dim arrD5(-1 To -1) As Variant: arrD5(-1) = Array()
    Dim arrD6(-1 To -1) As Variant: arrD6(-1) = Array(ArrToClct(Array()))
    Dim arrD7(-1 To -1) As Variant: arrD7(-1) = Array(ArrToClct(Array(0)))
    Dim arrD8(-3 To -1) As Variant
        arrD8(-3) = Array()
        arrD8(-2) = Array()
        arrD8(-1) = Array()
    Dim arrD9(-3 To -1) As Variant
        arrD9(-3) = Array()
        arrD9(-2) = Array()
        arrD9(-1) = Array(ArrToClct(Array()))

    Assert.AreEqual False, IsJagArr(arrD3), "D3" 'arr(-1)
    Assert.AreEqual False, IsJagArr(arrD4), "D4" 'arr(-1)
    Assert.AreEqual False, IsJagArr(arrD5), "D5" 'arr(-1)
    Assert.AreEqual False, IsJagArr(arrD6), "D6" 'arr(-1)(0) -> isODjeDt
    Assert.AreEqual False, IsJagArr(arrD7), "D7" 'arr(-1)(0) -> isODjeDt
    Assert.AreEqual False, IsJagArr(arrD8), "D8" 'arr(-3), arrD4(-2), arrD4(-1)
    Assert.AreEqual False, IsJagArr(arrD9), "D9" 'arr(-3)(0) -> isODjeDt

    'E: Dimention = 0
    Dim arrEInt(0) As Integer
    Dim arrELng(0) As Long
    Dim arrESng(0) As Single
    Dim arrEDbl(0) As Double
    Dim arrECur(0) As Currency
    Dim arrEByt(0) As Byte
    Dim arrEDat(0) As Date
    Dim arrEVar(0) As Variant
    
    Assert.AreEqual False, IsJagArr(arrEInt), "arrEInt"
    Assert.AreEqual False, IsJagArr(arrELng), "arrESng"
    Assert.AreEqual False, IsJagArr(arrESng), "arrESng"
    Assert.AreEqual False, IsJagArr(arrEDbl), "arrEDbl"
    Assert.AreEqual False, IsJagArr(arrECur), "arrECur"
    Assert.AreEqual False, IsJagArr(arrEByt), "arrEByt"
    Assert.AreEqual False, IsJagArr(arrEDat), "arrEDat"
    Assert.AreEqual False, IsJagArr(arrEVar), "arrEVar"
    
    Dim arrE0() As Variant: arrE0 = Array(ArrToClct(Array()))
    Dim arrE1() As Variant: arrE1 = Array(ArrToClct(Array(0)))
    
    Assert.AreEqual False, IsJagArr(arrE0), "arrE0"
    Assert.AreEqual False, IsJagArr(arrE1), "arrE1"


'   'F: Dimention > 0
    Dim arrFInt(0, 0) As Integer
    Dim arrFLng(0, 0) As Long
    Dim arrFSng(0, 0) As Single
    Dim arrFDbl(0, 0) As Double
    Dim arrFCur(0, 0) As Currency
    Dim arrFByt(0, 0) As Byte
    Dim arrFDat(0, 0) As Date
    Dim arrFVar(0, 0) As Variant

    Assert.AreEqual False, IsJagArr(arrFInt), "arrFInt"
    Assert.AreEqual False, IsJagArr(arrFLng), "arrFLng"
    Assert.AreEqual False, IsJagArr(arrFSng), "arrFSng"
    Assert.AreEqual False, IsJagArr(arrFDbl), "arrFDbl"
    Assert.AreEqual False, IsJagArr(arrFCur), "arrFCur"
    Assert.AreEqual False, IsJagArr(arrFByt), "arrFByt"
    Assert.AreEqual False, IsJagArr(arrFDat), "arrFDat"
    Assert.AreEqual False, IsJagArr(arrFVar), "arrFVar"
    
    Dim arrF0() As Variant: arrF0 = JagArrToArr2D(Array(Array(ArrToClct(Array())), Array(ArrToClct(Array()))))
    Dim arrF1() As Variant: arrF1 = JagArrToArr2D(Array(Array(ArrToClct(Array(0))), Array(ArrToClct(Array(0)))))

    Assert.AreEqual False, IsJagArr(arrF0), "arrF0"
    Assert.AreEqual False, IsJagArr(arrF1), "arrF1"

End Sub
