Attribute VB_Name = "Module1"
Sub VBAChallenge()
 ' Create Worksheet Variables
 Dim ws As Worksheet
 Dim lastRow As Long
 
 ' Variables we need to track
 Dim ticker As String
 Dim Open_Price As Double
 Dim Close_Price As Double
 Dim Total_Stock_Volume As Double
 Dim Quarterly_Change As Double
 Dim Percent_Change As Double
 
 ' Variables for summary tables
 Dim SummaryRow As Long
 Dim Greatest_Increase As Double
 Dim Greatest_Decrease As Double
 Dim Greatest_Total_Volume As Double
 Dim Greatest_Increase_Ticker As String
 Dim Greatest_Decrease_Ticker As String
 Dim Greatest_Total_Volume_Ticker As String
 
 ' Loop through each worksheet
 For Each ws In ThisWorkbook.Worksheets
    ' Set Inital Values for Variables
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    SummaryRow = 2
    Total_Stock_Volume = 0
    Greatest_Increase = 0
    Greatest_Decrease = 0
    Greatest_Total_Volume = 0
    
    'Table Headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Quarterly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    'Create Greatest Values Table
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
     
    ' Values for Ticker and Open Price
    ticker = ws.Cells(2, 1).Value
    Open_Price = ws.Cells(2, 3).Value
   
    ' Loop through Data
    For i = 2 To lastRow

    
        ' Formula to add to Stock Volume
        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
        
        ' Loop until last row of the current ticker
        If ws.Cells(i + 1, 1).Value <> ticker Then
            Close_Price = ws.Cells(i, 6).Value
            
            ' Calculate Changes for Percent Change and Quarterly Change
            Quarterly_Change = Close_Price - Open_Price
            If Open_Price <> 0 Then
                Percent_Change = Quarterly_Change / Open_Price
            Else
                Percent_Change = 0
            End If
            
            ' Transfer information to Summary Table
            ws.Cells(SummaryRow, 9).Value = ticker
            ws.Cells(SummaryRow, 10).Value = Quarterly_Change
            ws.Cells(SummaryRow, 11).Value = Percent_Change
            ws.Cells(SummaryRow, 12).Value = Total_Stock_Volume
            
            
    ' Conditional Format Summary Quarterly Change
        If Quarterly_Change > 0 Then
            ws.Cells(SummaryRow, 10).Interior.ColorIndex = 4
        ElseIf Quarterly_Change < 0 Then
            ws.Cells(SummaryRow, 10).Interior.ColorIndex = 3
        Else
            ws.Cells(SummaryRow, 10).Interior.ColorIndex = 0
        End If
        
    'Conditional Format Summary Percent Change
    If Percent_Change > 0 Then
            ws.Cells(SummaryRow, 11).Interior.ColorIndex = 4
        ElseIf Percent_Change < 0 Then
            ws.Cells(SummaryRow, 11).Interior.ColorIndex = 3
        Else
            ws.Cells(SummaryRow, 11).Interior.ColorIndex = 0
        End If
            
                ' Find Greatest Values
                If Percent_Change > Greatest_Increase Then
                    Greatest_Increase = Percent_Change
                    Greatest_Increase_Ticker = ticker
                End If
                
                If Percent_Change < Greatest_Decrease Then
                    Greatest_Decrease = Percent_Change
                    Greatest_Decrease_Ticker = ticker
                End If
                
                If Total_Stock_Volume > Greatest_Total_Volume Then
                    Greatest_Total_Volume = Total_Stock_Volume
                    Greatest_Total_Volume_Ticker = ticker
                End If
    
                
                ' Reset values of variables
                SummaryRow = SummaryRow + 1
                Total_Stock_Volume = 0
                ticker = ws.Cells(i + 1, 1).Value
                Open_Price = ws.Cells(i + 1, 3).Value
            End If
            
        Next i
        
        ' Fill in Greatest Values Summary Table
        ws.Cells(2, 16).Value = Greatest_Increase_Ticker
        ws.Cells(2, 17).Value = Greatest_Increase
        ws.Cells(3, 16).Value = Greatest_Decrease_Ticker
        ws.Cells(3, 17).Value = Greatest_Decrease
        ws.Cells(4, 16).Value = Greatest_Total_Volume_Ticker
        ws.Cells(4, 17).Value = Greatest_Total_Volume
        
        
        
    Next ws
    
      
 
End Sub
