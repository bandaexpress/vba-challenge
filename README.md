# vba-challenge
Assignment 2 from Columbia University Data Analytics Bootcamp

Attribute VB_Name = "Module1"
Sub StockMarketAnalysis()
    ' Define variables
    Dim LastRow As Long
    Dim Ticker As String
    Dim YearlyChange As Double
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim PercentageChange As Double
    Dim TotalVolume As Double
    Dim WorksheetName As String
    ' Define variables for tracking greatest metrics
    Dim MaxPercentageIncrease As Double
    Dim MinPercentageDecrease As Double
    Dim MaxTotalVolume As Double
    Dim MaxPercentageIncreaseTicker As String
    Dim MinPercentageDecreaseTicker As String
    Dim MaxTotalVolumeTicker As String
    ' Initialize the greatest metrics variables
    MaxPercentageIncrease = 0
    MinPercentageDecrease = 0
    MaxTotalVolume = 0
    
    ' Get the active worksheet
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Loop through all worksheets in the workbook
    For Each ws In ThisWorkbook.Sheets
       'Define the starting cell for output (e.g., Row 2, Column 9)
        Dim OutputRow As Long
        Dim OutputColumn As Long
        OutputRow = 2
        OutputColumn = 9
                
        ' Initialize variables for the current worksheet
        Ticker = ws.Cells(2, 1).Value
        OpeningPrice = ws.Cells(2, 3).Value
        TotalVolume = 0
        ' Find the last row of data in the current worksheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ' Loop through the data
        For i = 2 To LastRow
            ' Check if the Ticker symbol has changed
            If Ticker <> ws.Cells(i, 1).Value Then
                ' Calculate YearlyChange and PercentageChange
                ClosingPrice = ws.Cells(i - 1, 6).Value
                YearlyChange = ClosingPrice - OpeningPrice
                If OpeningPrice <> 0 Then
                    PercentageChange = (YearlyChange / OpeningPrice) * 100
                Else
                    PercentageChange = 0
                End If
                
                 ' Output the information to the specified cell
            ws.Cells(OutputRow, OutputColumn).Value = Ticker
            ws.Cells(OutputRow, OutputColumn + 1).Value = YearlyChange
            ws.Cells(OutputRow, OutputColumn + 2).Value = PercentageChange
            ws.Cells(OutputRow, OutputColumn + 3).Value = TotalVolume
            ' Apply conditional formatting based on YearlyChange
            If YearlyChange < 0 Then
                ' Set background color to red for negative changes
                ws.Cells(OutputRow, OutputColumn + 1).Interior.Color = RGB(255, 0, 0) ' Red
            ElseIf YearlyChange > 0 Then
                ' Set background color to green for positive changes
                ws.Cells(OutputRow, OutputColumn + 1).Interior.Color = RGB(0, 255, 0) ' Green
            End If
                ' Update greatest metrics
                If PercentageChange > MaxPercentageIncrease Then
                    MaxPercentageIncrease = PercentageChange
                    MaxPercentageIncreaseTicker = Ticker
                ElseIf PercentageChange < MinPercentageDecrease Then
                    MinPercentageDecrease = PercentageChange
                    MinPercentageDecreaseTicker = Ticker
                End If
                If TotalVolume > MaxTotalVolume Then
                    MaxTotalVolume = TotalVolume
                    MaxTotalVolumeTicker = Ticker
                End If
                ' Increment the output row
                OutputRow = OutputRow + 1
                ' Reset variables for the next Ticker
                Ticker = ws.Cells(i, 1).Value
                OpeningPrice = ws.Cells(i, 3).Value
                TotalVolume = 0
            End If
            ' Add to TotalVolume for the current Ticker
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        Next i
        
   ' Output the variable names as headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    ' Output the greatest metrics on the same worksheet in columns P and Q
    ws.Cells(1, 16).Value = "Greatest % Increase"
    ws.Cells(2, 16).Value = "Greatest % Decrease"
    ws.Cells(3, 16).Value = "Greatest Total Volume"
    
    ws.Cells(1, 17).Value = MaxPercentageIncreaseTicker
    ws.Cells(2, 17).Value = MinPercentageDecreaseTicker
    ws.Cells(3, 17).Value = MaxTotalVolumeTicker
    ' Add the corresponding metrics values
    ws.Cells(1, 18).Value = MaxPercentageIncrease & "%"
    ws.Cells(2, 18).Value = MinPercentageDecrease & "%"
    ws.Cells(3, 18).Value = MaxTotalVolume
    
    Next ws
End Sub
Sub SheetCount()
MsgBox Worksheets.Count
End Sub

"""
This Python script was created in part by ChatGPT, a language model developed by OpenAI.

Model: ChatGPT (GPT-3.5)
Source: OpenAI - https://openai.com
Date of Retrieval: 11/2/23
"""
