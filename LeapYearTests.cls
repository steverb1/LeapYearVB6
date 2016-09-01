VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LeapYearTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Sub test1997IsNotLeapYear()
    Dim LeapYearChecker As New LeapYearChecker
    Dim isLeapYear As Boolean
    isLeapYear = LeapYearChecker.isLeapYear(1997)
    Assert.IsFalse (isLeapYear)
End Sub

Sub test1996IsLeapYear()
    Dim LeapYearChecker As New LeapYearChecker
    Dim isLeapYear As Boolean
    isLeapYear = LeapYearChecker.isLeapYear(1996)
    Assert.IsTrue (isLeapYear)
End Sub

Sub test1900IsNotLeapYear()
    Dim LeapYearChecker As New LeapYearChecker
    Dim isLeapYear As Boolean
    isLeapYear = LeapYearChecker.isLeapYear(1900)
    Assert.IsFalse (isLeapYear)
End Sub

Sub test2000IsLeapYear()
    Dim LeapYearChecker As New LeapYearChecker
    Dim isLeapYear As Boolean
    isLeapYear = LeapYearChecker.isLeapYear(2000)
    Assert.IsTrue (isLeapYear)
End Sub
