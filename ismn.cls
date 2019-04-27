VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISMN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Base 1

Public Type ISMNInfo
  ISMN As String
  IDNumber As String
  CheckNumber As String
  Status As Boolean
End Type

Private mavarWeight As Variant
Private mlngLength As Long

Private Sub Class_Initialize()

  mavarWeight = Array(3, 1, 3, 1, 3, 1, 3, 1, 3)
  mlngLength = 10

End Sub

Public Function Check(ByVal ISMN As String) As Boolean

  Dim suma As Long
  Dim strISMN As String

  Check = False

  ISMN = Replace(ISMN, "-", "")
  strISMN = Replace(ISMN, "M", "3")
  If Len(ISMN) <> mlngLength Or Not IsNumeric(strISMN) Then Exit Function

  suma = CountSum(strISMN, mavarWeight)
  If Right(ISMN, 1) = CStr(10 - (suma Mod 10)) Then Check = True

End Function

Public Function GetInfo(ByVal ISMN As String) As ISMNInfo

  Dim Info As ISMNInfo

  With Info
    .ISMN = ISMN
    .Status = Check(ISMN)
    If .Status Then
      ISMN = Replace(ISMN, "-", "")
      .IDNumber = Mid(ISMN, 1, 9)
      .CheckNumber = Right(ISMN, 1)
    End If
  End With
  GetInfo = Info

End Function
