Attribute VB_Name = "Module1"
' Homa Nabili
' Data Analytics Winter 2020
' Homework #2
' Feb. 8, 2010

Sub StockYearlyChangeAllSheets()

    Dim row As Long
    Dim rowCount As Long
    Dim i As Long
    
    'stock's start value at begining of the year
    Dim StartValue As Double
    'stock's end value at end of year
    Dim EndValue As Double
    'stock's total volume
    Dim totalVolume As Double
    'boolean variable to keep track of first occurance of a stock symbol in the list
    Dim Found As Boolean
    'object for each worksheet
    Dim ws As Worksheet
    
    'loop through all the worksheets in current workbook
    For Each ws In Worksheets
        ws.Activate
        'count rows in current sheet
        rowCount = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
        'initialize variables
        i = 2 'skip header line and start at row 2
        totalVolume = 0
        Found = False
        
        'add titles
        ws.Range("J1").Value = "Ticker"
        ws.Range("K1").Value = "Yearly Change"
        ws.Range("L1").Value = "Percent Change"
        ws.Range("M1").Value = "Total Stock Volume"
    
        For row = 2 To rowCount
            totalVolume = totalVolume + ws.Cells(row, 7).Value
            If Found = False Then
                'read opening price in the first line the ticker appears
                StartValue = ws.Cells(row, 3).Value
                Found = True
            End If
            'if next line is start of a new ticker, read End Price of current ticker, calculate and save values
            If ws.Cells(row, 1).Value <> ws.Cells(row + 1, 1).Value Then
                EndValue = ws.Cells(row, 6).Value
                ws.Cells(i, 10).Value = ws.Cells(row, 1).Value
                ws.Cells(i, 11).Value = EndValue - StartValue
                If StartValue <> 0 Then
                    ws.Cells(i, 12).Value = Round((EndValue - StartValue) / StartValue, 2)
                Else
                    ws.Cells(i, 12).Value = ""
                End If
                
                ws.Cells(i, 13).Value = totalVolume
                i = i + 1
                'reset total volume for next ticker
                totalVolume = 0
                'reset boolean for next ticker
                Found = False
            End If
        
        Next row

        'set conditional formatting for Yearly Change
        SetFormatConditions (i)

        ws.Range("O1").Value = "Greatest % increase"
        ws.Range("O2").Value = "Greatest % decrease"
        ws.Range("O3").Value = "Greatest total volume"

        'Find where result rows end
        LastRow = ws.Cells(ActiveSheet.Rows.Count, "L").End(xlUp).row

        ws.Range("P1").Value = "=MAX(L2:L" & LastRow & ")" 'Greatest % increase
        ws.Range("P2").Value = "=MIN(L2:L" & LastRow & ")" 'Greatest % decrease
        ws.Range("P3").Value = "=MAX(M2:M" & LastRow & ")" 'Greatest total volume
        
        ' Autofit to display data
        ws.Columns("J:P").AutoFit
    
    'go to next worksheet
    Next ws

End Sub


Sub StockYearlyChange()

    Dim row As Long
    Dim rowCount As Long
    Dim i As Long
    
    'stock's start value at begining of the year
    Dim StartValue As Double
    'stock's end value at end of year
    Dim EndValue As Double
    'stock's total volume
    Dim totalVolume As Double
    'boolean variable to keep track of first occurance of a stock symbol in the list
    Dim Found As Boolean
    
    Found = False
    
    'add titles
    ActiveSheet.Range("J1").Value = "Ticker"
    ActiveSheet.Range("K1").Value = "Yearly Change"
    ActiveSheet.Range("L1").Value = "Percent Change"
    ActiveSheet.Range("M1").Value = "Total Stock Volume"
        
    rowCount = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).row
    i = 2
    totalVolume = 0

    For row = 2 To rowCount
        totalVolume = totalVolume + ActiveSheet.Cells(row, 7).Value
        If Found = False Then
            'read opening price in the first line the ticker appears
            StartValue = ActiveSheet.Cells(row, 3).Value
            Found = True
        End If
        'if next line is start of a new ticker, read End Price of current ticker, calculate and save values
        If ActiveSheet.Cells(row, 1).Value <> ActiveSheet.Cells(row + 1, 1).Value Then
            EndValue = ActiveSheet.Cells(row, 6).Value
            ActiveSheet.Cells(i, 10).Value = ActiveSheet.Cells(row, 1).Value
            ActiveSheet.Cells(i, 11).Value = EndValue - StartValue
            ActiveSheet.Cells(i, 12).Value = Round((EndValue - StartValue) / StartValue, 2)
            ActiveSheet.Cells(i, 13).Value = totalVolume
            i = i + 1
            'reset total volume for next ticker
            totalVolume = 0
            'reset boolean for next ticker
            Found = False
        End If
        
    Next row

    'set conditional formatting for Yearly Change
    SetFormatConditions (i)

    ActiveSheet.Range("O1").Value = "Greatest % increase"
    ActiveSheet.Range("O2").Value = "Greatest % decrease"
    ActiveSheet.Range("O3").Value = "Greatest total volume"

    'Find where result rows end
    LastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, "L").End(xlUp).row

    ActiveSheet.Range("P1").Value = "=MAX(L2:L" & LastRow & ")" 'Greatest % increase
    ActiveSheet.Range("P2").Value = "=MIN(L2:L" & LastRow & ")" 'Greatest % decrease
    ActiveSheet.Range("P3").Value = "=MAX(M2:M" & LastRow & ")" 'Greatest total volume
        
    ' Autofit to display data
    ActiveSheet.Columns("J:P").AutoFit
        
End Sub


Sub SetFormatConditions(LastRow As Long)

    ActiveSheet.Range("K2:K" & LastRow).Select

    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399945066682943
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False

End Sub


