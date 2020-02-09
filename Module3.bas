Attribute VB_Name = "Module3"
Sub Stock_results()

 Dim Stock_name As String
 
 Dim Stock_Total As Double
 Stock_Total = 0
 
 Dim Yearly_change As Double
 Yearly_change = 0
 
 Dim Percentage_change As Double
 Percentage_change = 0
 
 
 open_stock_price = Cells(2, 3).Value
 close_stock_price = Cells(2, 6).Value
 
 
 
 Dim Summary_Table_Row As Integer
 Summary_Table_Row = 2
 
 For i = 2 To 705714
 
 If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
 Stock_name = Cells(i, 1).Value
 Stock_Total = Stock_Total + Cells(i, 7).Value
 close_stock_price = Cells(i, 6).Value
 Yearly_change = close_stock_price - open_stock_price
 Percentage_change = Yearly_change / open_stock_price

 

 
 Range("I" & Summary_Table_Row).Value = Stock_name
 Range("L" & Summary_Table_Row).Value = Stock_Total
 Range("J" & Summary_Table_Row).Value = Yearly_change
 
 Range("K" & Summary_Table_Row).Value = Percentage_change
 
Summary_Table_Row = Summary_Table_Row + 1
 Stock_Total = 0
 open_stock_price = Cells(i + 1, 3).Value
 
 
 Else
 Stock_Total = Stock_Total + Cells(i, 7).Value
 
 If Range("J").Value = "+" Then
 Range("J2:J2836").Interior.ColorIndex = 4
 
 End If
 End If
 Next i
 End Sub

