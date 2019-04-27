VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdISSN 
      Caption         =   "Info"
      Height          =   615
      Index           =   1
      Left            =   3480
      TabIndex        =   23
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdISSN 
      Caption         =   "Sprawdü"
      Height          =   615
      Index           =   0
      Left            =   2400
      TabIndex        =   22
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox txtISSN 
      Height          =   285
      Left            =   120
      TabIndex        =   21
      Text            =   "0867-0153"
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CommandButton cmdISMN 
      Caption         =   "Info"
      Height          =   615
      Index           =   1
      Left            =   3480
      TabIndex        =   19
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdISMN 
      Caption         =   "Sprawdü"
      Height          =   615
      Index           =   0
      Left            =   2400
      TabIndex        =   18
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox txtISMN 
      Height          =   285
      Left            =   120
      TabIndex        =   17
      Text            =   "M9005202-2-7"
      Top             =   3840
      Width           =   2175
   End
   Begin VB.CommandButton cmdISBN 
      Caption         =   "Info"
      Height          =   615
      Index           =   1
      Left            =   3480
      TabIndex        =   15
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdISBN 
      Caption         =   "Sprawdü"
      Height          =   615
      Index           =   0
      Left            =   2400
      TabIndex        =   14
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox txtISBN 
      Height          =   285
      Left            =   120
      TabIndex        =   13
      Text            =   "83-85784-25-X"
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CommandButton cmdREGON 
      Caption         =   "Info"
      Height          =   615
      Index           =   1
      Left            =   3480
      TabIndex        =   11
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdREGON 
      Caption         =   "Sprawdü"
      Height          =   615
      Index           =   0
      Left            =   2400
      TabIndex        =   10
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtREGON 
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Text            =   "590096454"
      Top             =   2160
      Width           =   2175
   End
   Begin VB.CommandButton cmdNIP 
      Caption         =   "Info"
      Height          =   615
      Index           =   1
      Left            =   3480
      TabIndex        =   7
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdNIP 
      Caption         =   "Sprawdü"
      Height          =   615
      Index           =   0
      Left            =   2400
      TabIndex        =   6
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtNIP 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Text            =   "625-215-87-49"
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton cmdPESEL 
      Caption         =   "Info"
      Height          =   615
      Index           =   1
      Left            =   3480
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdPESEL 
      Caption         =   "Sprawdü"
      Height          =   615
      Index           =   0
      Left            =   2400
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtPESEL 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "77021605871"
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label lblISSN 
      Caption         =   "ISSN :"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label lblISMN 
      Caption         =   "ISMN :"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   120
      X2              =   4560
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   120
      X2              =   4560
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label lblISBN 
      Caption         =   "ISBN :"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   120
      X2              =   4560
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label lblREGON 
      Caption         =   "REGON :"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   120
      X2              =   4560
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label lnlNIP 
      Caption         =   "NIP :"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   2175
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   120
      X2              =   4560
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lblPESEL 
      Caption         =   "PESEL :"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdISBN_Click(Index As Integer)

  Dim v As New Validator.ISBN
  Dim i As Validator.ISBNInfo
  Dim strCountry As String

  Set v = New Validator.ISBN
  If Index = 0 Then
    If v.Check(txtISBN.Text) Then
      MsgBox "Status: OK", , "ISBN"
    Else
      MsgBox "Status: B≥πd", , "ISBN"
    End If
  Else
    i = v.GetInfo(txtISBN.Text)
    Select Case i.Country
      Case CountryEnum.unknown: strCountry = "nieznany"
      Case CountryEnum.Poland:  strCountry = "Polska"
      Case Else:                strCountry = "?"
    End Select
    MsgBox _
      "ISBN: " & i.ISBN & vbCrLf & _
      "Status: " & IIf(i.Status, "OK", "B≥πd") & vbCrLf & _
      "Kraj wydania: " & strCountry & vbCrLf & _
      "Numer: " & i.IDNumber & vbCrLf & _
      "Cyfra kontrolna: " & i.CheckNumber _
      , , "ISBN"
  End If

End Sub


Private Sub cmdISMN_Click(Index As Integer)

  Dim v As New Validator.ISMN
  Dim i As Validator.ISMNInfo

  Set v = New Validator.ISMN
  If Index = 0 Then
    If v.Check(txtISMN.Text) Then
      MsgBox "Status: OK", , "ISMN"
    Else
      MsgBox "Status: B≥πd", , "ISMN"
    End If
  Else
    i = v.GetInfo(txtISMN.Text)
    MsgBox _
      "ISMN: " & i.ISMN & vbCrLf & _
      "Status: " & IIf(i.Status, "OK", "B≥πd") & vbCrLf & _
      "Numer: " & i.IDNumber & vbCrLf & _
      "Cyfra kontrolna: " & i.CheckNumber _
      , , "ISMN"
  End If

End Sub

Private Sub cmdISSN_Click(Index As Integer)

  Dim v As New Validator.ISSN
  Dim i As Validator.ISSNInfo

  Set v = New Validator.ISSN
  If Index = 0 Then
    If v.Check(txtISSN.Text) Then
      MsgBox "Status: OK", , "ISSN"
    Else
      MsgBox "Status: B≥πd", , "ISSN"
    End If
  Else
    i = v.GetInfo(txtISSN.Text)
    MsgBox _
      "ISSN: " & i.ISSN & vbCrLf & _
      "Status: " & IIf(i.Status, "OK", "B≥πd") & vbCrLf & _
      "Numer: " & i.IDNumber & vbCrLf & _
      "Cyfra kontrolna: " & i.CheckNumber _
      , , "ISSN"
  End If

End Sub


Private Sub cmdNIP_Click(Index As Integer)

  Dim v As New Validator.NIP
  Dim i As Validator.NIPInfo

  If Index = 0 Then
    If v.Check(txtNIP.Text) Then
      MsgBox "Status: OK", , "NIP"
    Else
      MsgBox "Status: B≥πd", , "NIP"
    End If
  Else
    i = v.GetInfo(txtNIP.Text)
    MsgBox _
      "NIP: " & i.NIP & vbCrLf & _
      "Status: " & IIf(i.Status, "OK", "B≥πd") & vbCrLf & _
      "Numer: " & i.IDNumber & vbCrLf & _
      "Cyfra kontrolna: " & i.CheckNumber _
      , , "NIP"
  End If

End Sub

Private Sub cmdPESEL_Click(Index As Integer)

  Dim v As New Validator.PESEL
  Dim i As Validator.PESELInfo

  If Index = 0 Then
    If v.Check(txtPESEL.Text) Then
      MsgBox "Status: OK", , "PESEL"
    Else
      MsgBox "Status: B≥πd", , "PESEL"
    End If
  Else
    i = v.GetInfo(txtPESEL.Text)
    MsgBox _
      "PESEL: " & i.PESEL & vbCrLf & _
      "Status: " & IIf(i.Status, "OK", "B≥πd") & vbCrLf & _
      "Rok urodzenia: " & i.BirthDate & vbCrLf & _
      "Numer: " & i.IDNumber & vbCrLf & _
      "P≥eÊ: " & Choose(1 + i.Sex, "Kobieta", "MÍøczyzna") & vbCrLf & _
      "Cyfra kontrolna: " & i.CheckNumber _
      , , "PESEL"
  End If

End Sub


Private Sub cmdREGON_Click(Index As Integer)

  Dim v As New Validator.REGON
  Dim i As Validator.REGONInfo

  If Index = 0 Then
    If v.Check(txtREGON.Text) Then
      MsgBox "Status: OK", , "REGON"
    Else
      MsgBox "Status: B≥πd", , "REGON"
    End If
  Else
    i = v.GetInfo(txtREGON.Text)
    MsgBox _
      "REGON: " & i.REGON & vbCrLf & _
      "Status: " & IIf(i.Status, "OK", "B≥πd") & vbCrLf & _
      "WojewÛdztwo: " & i.Wojewodztwo & vbCrLf & _
      "Numer: " & i.IDNumber & vbCrLf & _
      "Cyfra kontrolna: " & i.CheckNumber _
      , , "REGON"
  End If

End Sub


