Attribute VB_Name = "Module1"
Sub EasyModerate()
    Dim Ticker As String
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    MsgBox (LastRow)
    Dim Total_Volume As Double
    Total_Volume = 0
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    Dim Close_Price As Double
    Dim Open_Price As Double
For i = 2 To LastRow
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Ticker = Cells(i, 1).Value
        Total_Volume = Total_Volume + Cells(i, 7).Value
        Range("J" & Summary_Table_Row).Value = Ticker
        Range("K" & Summary_Table_Row).Value = Total_Volume
        'Summary_Table_Row = Summary_Table_Row + 1
        Total_Volume = 0
    Else
        Total_Volume = Total_Volume + Cells(i, 7).Value
    End If
    If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
        Open_Price = Cells(i, 3).Value
    ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Close_Price = Cells(i, 6).Value
        Range("L" & Summary_Table_Row).Value = Close_Price - Open_Price
    If Open_Price = 0 Then
        Range("M" & Summary_Table_Row).Value = "N/A"
        Summary_Table_Row = Summary_Table_Row + 1
    Else
        Range("M" & Summary_Table_Row).Value = Format((Close_Price - Open_Price) / Open_Price, "Percent")
        Summary_Table_Row = Summary_Table_Row + 1
    End If
    End If
Next i
For i = 2 To LastRow
    If Cells(i, 12) > 0 Then
        Cells(i, 12).Interior.ColorIndex = 4
    ElseIf Cells(i, 12) < 0 Then
        Cells(i, 12).Interior.ColorIndex = 3
    End If
Next i
End Sub
