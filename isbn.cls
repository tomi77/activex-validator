VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISBN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Base 1

Public Enum CountryEnum
  unknown = -1
  Poland = 83
End Enum

Public Type ISBNInfo
  ISBN As String
  Country As CountryEnum
  IDNumber As String
  CheckNumber As String
  Status As Boolean
End Type

Private mavarWeight As Variant
Private mlngLength As Long

Private Sub Class_Initialize()

  mavarWeight = Array(10, 9, 8, 7, 6, 5, 4, 3, 2)
  mlngLength = 10

End Sub

Public Function Check(ByVal ISBN As String) As Boolean

  Dim suma As Long
  Dim ControlChar As String

  Check = False

  ISBN = Replace(ISBN, "-", "")
  If Len(ISBN) <> mlngLength Or Not IsNumeric(Left(ISBN, mlngLength - 1)) Then Exit Function

  suma = CountSum(ISBN, mavarWeight)
  Select Case 11 - (suma Mod 11)
    Case 11: ControlChar = "0"
    Case 10: ControlChar = "X"
    Case Else: ControlChar = CStr(11 - (suma Mod 11))
  End Select
  If Right(ISBN, 1) = ControlChar Then Check = True

End Function

Public Function GetInfo(ByVal ISBN As String) As ISBNInfo

  Dim Info As ISBNInfo
  Dim lngCountryPos As Long

  With Info
    .ISBN = ISBN
    .Status = Check(ISBN)
    If .Status Then
      If Mid(ISBN, 2, 1) = "-" Or Mid(ISBN, 3, 1) = "-" Then
        lngCountryPos = InStr(ISBN, "-")
        .Country = Left(ISBN, lngCountryPos - 1)
      Else
        lngCountryPos = 1
        .Country = unknown
      End If
      ISBN = Replace(ISBN, "-", "")
      .IDNumber = Mid(ISBN, lngCountryPos, mlngLength - 1 - lngCountryPos + 1)
      .CheckNumber = Right(ISBN, 1)
    End If
  End With
  GetInfo = Info

End Function

