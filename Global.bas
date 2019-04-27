Attribute VB_Name = "basGlobal"
Option Explicit
Option Base 1


Public Function CountSum(Number As String, Weights As Variant) As Long

  Dim i As Long
  Dim lngUBound As Long
  Dim lngSum As Long

  lngSum = 0
  lngUBound = UBound(Weights)

  For i = 1 To lngUBound
    lngSum = lngSum + CLng(Mid(Number, i, 1)) * Weights(i)
  Next i
  CountSum = lngSum

End Function
