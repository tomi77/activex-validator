VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Base 1

Public Type NIPInfo
  NIP As String
  IDNumber As String
  CheckNumber As Long
  Status As Boolean
End Type

Private mavarWeight As Variant
Private mlngLength As Long

Private Sub Class_Initialize()

  mavarWeight = Array(6, 5, 7, 2, 3, 4, 5, 6, 7)
  mlngLength = 10

End Sub

Public Function Check(ByVal NIP As String) As Boolean

  Dim suma As Long

  Check = False

  NIP = Replace(NIP, "-", "")
  If Len(NIP) <> mlngLength Or Not IsNumeric(NIP) Then Exit Function

  suma = CountSum(NIP, mavarWeight)
  suma = suma Mod 11
  If Right(NIP, 1) = Mid(suma, Len(CStr(suma)), 1) Then Check = True

End Function

Public Function GetInfo(ByVal NIP As String) As NIPInfo

  Dim Info As NIPInfo

  With Info
    .NIP = NIP
    .Status = Check(NIP)
    If .Status Then
      NIP = Replace(NIP, "-", "")
      .IDNumber = Mid(NIP, 1, 9)
      .CheckNumber = Right(NIP, 1)
    End If
  End With
  GetInfo = Info

End Function
