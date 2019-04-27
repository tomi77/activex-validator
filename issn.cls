VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISSN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 1

Public Type ISSNInfo
  ISSN As String
  IDNumber As String
  CheckNumber As String
  Status As Boolean
End Type

Private mavarWeight As Variant
Private mlngLength As Long

Private Sub Class_Initialize()

  mavarWeight = Array(8, 7, 6, 5, 4, 3, 2)
  mlngLength = 8

End Sub

Public Function Check(ByVal ISSN As String) As Boolean

  Dim suma As Long
  Dim ControlChar As String

  Check = False

  ISSN = Replace(ISSN, "-", "")
  If Len(ISSN) <> mlngLength Or Not IsNumeric(Left(ISSN, mlngLength - 1)) Then Exit Function

  suma = CountSum(ISSN, mavarWeight)
  Select Case 11 - (suma Mod 11)
    Case 11: ControlChar = "0"
    Case 10: ControlChar = "X"
    Case Else: ControlChar = CStr(11 - (suma Mod 11))
  End Select
  If Right(ISSN, 1) = ControlChar Then Check = True

End Function

Public Function GetInfo(ByVal ISSN As String) As ISSNInfo

  Dim Info As ISSNInfo

  With Info
    .ISSN = ISSN
    .Status = Check(ISSN)
    If .Status Then
      ISSN = Replace(ISSN, "-", "")
      .IDNumber = Mid(ISSN, 1, 7)
      .CheckNumber = Right(ISSN, 1)
    End If
  End With
  GetInfo = Info

End Function


