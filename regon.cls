VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "REGON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Base 1

Public Type REGONInfo
  REGON As String
  Wojewodztwo As String
  IDNumber As String
  CheckNumber As Long
  Status As Boolean
End Type

Private mavarWeight As Variant
Private mlngLength As Long

Private Sub Class_Initialize()

  mavarWeight = Array(8, 9, 2, 3, 4, 5, 6, 7)
  mlngLength = 9

End Sub

Public Function Check(ByVal REGON As String) As Boolean

  Dim suma As Long

  Check = False

  If Len(REGON) <> mlngLength Or Not IsNumeric(REGON) Then Exit Function

  suma = CountSum(REGON, mavarWeight)
  suma = suma Mod 11
  If Right(REGON, 1) = Mid(suma, Len(CStr(suma)), 1) Then Check = True

End Function

Public Function GetInfo(ByVal REGON As String) As REGONInfo

  Dim Info As REGONInfo

  With Info
    .REGON = REGON
    .Status = Check(REGON)
    If .Status Then
      .Wojewodztwo = Mid(REGON, 1, 2)
      .IDNumber = Mid(REGON, 3, 6)
      .CheckNumber = Right(REGON, 1)
    End If
  End With
  GetInfo = Info

End Function

