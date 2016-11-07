

Dim Coefficeint() As Double
Dim Order As Double
Dim Row As Double
Dim Column As Double
Dim Total As Double
Dim List As Collection
Dim CArray As Variant
Dim DArray As Variant


Sub Calculate()
    ' actual code to print check
    
    CArray = Sheets("Sheet").Range("C10:C15").Value
    DArray = Sheets("Sheet").Range("D10:D13").Value
    Order = 5
    Dim root As Double
    For Each Item In arr
        Row = Item
        For i = 1 To Order
        root = Item ^ i
        Row = Row + Coefficient(i) * root
        Next
        Total = Total + Row
    Next Item
    
End Sub


Sub Button1_Click()
Calculate()
End Sub
