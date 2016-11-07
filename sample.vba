Option Explicit

Public Type PolyType
a() As Double
xvalue As Double
Order As Integer
poly As Double
deriv As Double
x As Double

End Type

Sub Button1_Click()

Dim ValueRange As Range
Dim Row As Double
Dim Order As Double
Dim Column As Double
Dim Total As Double
Dim List As Collection
Dim CArray As Variant
Dim root As Double
Dim Item As Range
Dim i As Integer

  ' actual code to print check
    CArray = Array(1, 2, 3, 4, 5)
    ValueRange = Sheet1.Range("C10:C100")
    Order = 5
   
    For Each Item In Sheet1.Range("C10:C100")
        Row = Item
        For i = 1 To Order
        root = Item ^ i
        Row = Row + ValueRange(i) * root
        Next
        List.Add (Row)
    Next Item
End Sub
