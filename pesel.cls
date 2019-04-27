VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PESEL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Base 1

Public Enum SexEnum
  Female = 0
  Male = 1
End Enum

Public Type PESELInfo
  PESEL As String
  BirthDate As Date
  IDNumber As String
  Sex As SexEnum
  CheckNumber As Long
  Status As Boolean
End Type

Private mavarWeight As Variant
Private mlngLength As Long

Private Sub Class_Initialize()

  mavarWeight = Array(1, 3, 7, 9, 1, 3, 7, 9, 1, 3)
  mlngLength = 11

End Sub

Public Function Check(ByVal PESEL As String) As Boolean

  Dim suma As Long

  Check = False

  If Len(PESEL) <> mlngLength Or Not IsNumeric(PESEL) Then Exit Function

  suma = CountSum(PESEL, mavarWeight)
  If CLng(Right(PESEL, 1)) = 10 - (suma Mod 10) Then Check = True

End Function

Public Function GetInfo(ByVal PESEL As String) As PESELInfo

  Dim Info As PESELInfo

  With Info
    .PESEL = PESEL
    .Status = Check(PESEL)
    If .Status Then
      .BirthDate = _
        Choose(CLng(Mid(PESEL, 3, 2)) \ 20 + 1, "19", "20", "00", "00", "18") & Mid(PESEL, 1, 2) & _
        "-" & Right("0" & CStr(CLng(Mid(PESEL, 3, 2)) Mod 20), 2) & _
        "-" & Mid(PESEL, 5, 2)
      .IDNumber = Mid(PESEL, 7, 3)
      .Sex = CLng(Mid(PESEL, 10, 1)) Mod 2
      .CheckNumber = Right(PESEL, 1)
    End If
  End With
  GetInfo = Info

End Function
