Attribute VB_Name = "EB_ExportExcelFile2"
Sub FormalTemplate()
Dim SectorNum As Integer

For SectorNum = 1 To SectorMaxNum

    ActiveWorkbook.Sheets.Add(After:=Worksheets(ActiveWorkbook.Worksheets.Count)).name = "Sector " & SectorNum
    
    ActiveSheet.Cells.Clear

    Call InsertData(SectorNum)
    Call TitleBar
    Call RSRPLossTable
    Call ComponentLoss
    Call AddEquation
    Call ColoringAndBorder
    
    SectorRowCount = 0
    
Next
End Sub

Sub InsertData(SectorNum As Integer)


    SectorRowCount = 1
    
    For i = 1 To AntCount
        If MaterialList(i, ArrSector) = "Sector " & SectorNum Then
            For j = 1 To ArrJumper
                Cells(8 + SectorRowCount, j).Value = MaterialList(i, j)
            Next

            For j = Arr2WaySplitter To ArrCombiner
                Cells(8 + SectorRowCount, 1 + j).Value = MaterialList(i, j)
            Next
            
            Cells(8 + SectorRowCount, 25).Value = MaterialList(i, ArrAntGain)
            
            SectorRowCount = SectorRowCount + 1
        End If
    Next
    SectorRowCount = SectorRowCount - 1

    With ActiveSheet.Sort

        .SortFields.Add Key:=Range("A8"), Order:=xlAscending
        .SetRange Range("A8:Y" & 8 + SectorRowCount)
        .Header = xlYes
        .Apply

    End With

    'Antenna Type
    For p = 1 To SectorRowCount
        If InStr(Cells(p + 8, ArrFloor).Value, "L") And Cells(p + 8, ArrFloor).Value <> "LG" Then
            Cells(p + 8, 2).Value = "AL"
        Else
            Cells(p + 8, 2).Value = "A"
        End If
    Next
    
    'Sector
    Range(Cells(9, 1), Cells(8 + SectorRowCount, 1)) = 1

End Sub

Sub AddEquation()

        'Cable Loss
        Range("I9").Formula = Format("=SUMPRODUCT(E9:H9,$E$6:$H$6)", "0.000")
        'Device Loss
        Range("V9").Formula = "=SUMPRODUCT(J9:U9,$J$6:$U$6)"
        'Total Path Loss
        Range("W9").Formula = "=I9+V9"
        'BTS/RU Output
        Range("X9").Formula = "=$Z$1"
        'EIRP
        Range("Z9").Formula = "=X9-W9+Y9"
        'RSRP
        Range("AA9").Formula = "=IF($B9=""A"", Z9-$Z$2, IF($B9=""AL"", Z9-$Z$3, FALSE))"
        'Pass/Fail
        Range("AB9").Formula = "=IF(AND($Z$4=3500, $AA9>=-104), ""PASS"", IF(AND($Z$4=2600, $AA9>=-95), ""PASS"", ""FAIL""))"
        
        If SectorRowCount > 1 Then
            Range("I9").AutoFill Destination:=Range("I9:I" & 8 + SectorRowCount), Type:=xlFillCopy
            Range("V9").AutoFill Destination:=Range("V9:V" & 8 + SectorRowCount), Type:=xlFillCopy
            Range("W9").AutoFill Destination:=Range("W9:W" & 8 + SectorRowCount), Type:=xlFillCopy
            Range("X9").AutoFill Destination:=Range("X9:X" & 8 + SectorRowCount), Type:=xlFillCopy
            Range("Z9").AutoFill Destination:=Range("Z9:Z" & 8 + SectorRowCount), Type:=xlFillCopy
            Range("AA9").AutoFill Destination:=Range("AA9:AA" & 8 + SectorRowCount), Type:=xlFillCopy
            Range("AB9").AutoFill Destination:=Range("AB9:AB" & 8 + SectorRowCount), Type:=xlFillCopy
        End If
        

        

End Sub

Sub RSRPLossTable()

    Range("Y1").Value = "RU Output RSRP (dBm)"
    Range("Z1").Value = RSRP_output
    Range("AA1").Value = "dBm"
    
    Range("Y2").Value = "FSPL(Indoor)"
    Range("Z2").Value = FSPL
    Range("AA2").Value = "dB"
    
    Range("Y3").Value = "FSPL(Lift)"
    Range("Z3").Value = FSPL_lift
    Range("AA3").Value = "dB"

    Range("Y4").Value = "Freq."
    Range("Z4").Value = FreqChoice

End Sub

Sub ComponentLoss()

    'Cable Loss Title
    Range("E5").Value = "Cable Loss:"
    Range("E5:G5").Merge
    Range("E5:G5").HorizontalAlignment = xlCenter
    
    'Cable
    Range("E6").FormulaR1C1 = "=IF(R4C26=3500, 0.147, IF(R4C26=2600, 12.4/100, FALSE))"
    Range("F6").FormulaR1C1 = "=IF(R4C26=3500, 0.0795, IF(R4C26=2600, 6.53/100, FALSE))"
    Range("G6").FormulaR1C1 = "=IF(R4C26=3500, 0.058, IF(R4C26=2600, 4.9/100, FALSE))"
    'Jumper
    Range("H6").Value = 0.5
    'C3
    Range("J6").Value = 3.6
    Range("K6").Value = 5.6
    'C6
    Range("L6").Value = 1.7
    Range("M6").Value = 7
    'C10
    Range("N6").Value = 1
    Range("O6").Value = 11.3
    'C15
    Range("P6").Value = 0.5
    Range("Q6").Value = 16.3
    'C20
    Range("R6").Value = 0.2
    Range("S6").Value = 21.3
    'Hybrid & Combiner
    Range("T6").Value = 3.1
    Range("U6").Value = 1

    
    
    

End Sub

Sub TitleBar()


    Range("A8").Value = "Sector"
    Range("B8").Value = "Antenna"
    Range("C8").Value = "Floor"

    Range("D7").Value = "Antenna"
    Range("D7:D8").Merge

    Range("E7").Value = "LCF12"
    Range("F7").Value = "LCF78"
    Range("G7").Value = "LCF114"
    Range("E8").Value = "Length(m)"

    Range("H7").Value = "Jumper"
    Range("H8").Value = "pcs"

    Range("I7").Value = "Cable" & vbCrLf & "Loss"
    Range("I8").Value = "dB"

    Range("J7").Value = "2-way" & vbCrLf & "Splitter"
    Range("K7").Value = "3-way" & vbCrLf & "Splitter"

    Range("L7").Value = "6dB Coupler"
    Range("L8").Value = "Thr."
    Range("M8").Value = "Couple"

    Range("N7").Value = "10dB Coupler"
    Range("N8").Value = "Thr."
    Range("O8").Value = "Couple"

    Range("P7").Value = "15dB Coupler"
    Range("P8").Value = "Thr."
    Range("Q8").Value = "Couple"

    Range("R7").Value = "20dB Coupler"
    Range("R8").Value = "Thr."
    Range("S8").Value = "Couple"

    Range("T7").Value = "Hybrid"
    Range("U7").Value = "QBC"
    Range("V7").Value = "Device" & vbCrLf & "Loss"
    Range("W7").Value = "Total" & vbCrLf & "Path" & vbCrLf & "Loss"
    Range("X7").Value = "BTS/" & vbCrLf & "RU" & vbCrLf & "Output"

    Range("Y7").Value = "Antenna" & vbCrLf & "Gain"
    Range("Z7").Value = "EIRP"
    Range("AA7").Value = "RSRP"

    Range("AB7").Value = "PASS/FAIL"
    Range("AB8").Value = "Band 2600 >= -95 dBm" & vbCrLf & "Band 3500 >= -104 dBm"

    '________________________________________________________

    Range("D7:D8").Merge
    Range("E8:G8").Merge
    Range("J7:J8").Merge
    Range("K7:K8").Merge
    Range("L7:M7").Merge
    Range("N7:O7").Merge
    Range("P7:Q7").Merge
    Range("R7:S7").Merge
    Range("T7:T8").Merge
    Range("U7:U8").Merge
    Range("V7:V8").Merge
    Range("W7:W8").Merge
    Range("X7:X8").Merge
    Range("Y7:Y8").Merge
    Range("Y7:Y8").Merge
    Range("Z7:Z8").Merge
    Range("AA7:AA8").Merge

    Columns("A:AA").ColumnWidth = 8
    Columns("AB").ColumnWidth = 22

    With Range("D7:AB8").Borders
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With

    With Range("D7:AB8").Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With

    With Range("D7:AB8").Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With

    Range("D7:AB8").HorizontalAlignment = xlCenter
    Range("A8:C8").HorizontalAlignment = xlCenter

    Range("D7:AB8").VerticalAlignment = xlCenter
    Range("A8:C8").VerticalAlignment = xlCenter

End Sub

Sub ColoringAndBorder()

Dim FRow As Integer
Dim MergeRow As Integer
Dim MergeState As Integer

    'Equations
    Range(Cells(9, 22), Cells(8 + SectorRowCount, 24)).Interior.Color = RGB(238, 236, 225)
    Range(Cells(9, 26), Cells(8 + SectorRowCount, 27)).Interior.Color = RGB(238, 236, 225)
    
    With Range("AB9:AB" & 8 + SectorRowCount).FormatConditions.Add(xlCellValue, xlEqual, "PASS")
        .Interior.Color = RGB(153, 255, 102)
    End With
    
    With Range("AB9:AB" & 8 + SectorRowCount).FormatConditions.Add(xlCellValue, xlEqual, "FAIL")
        .Interior.Color = RGB(255, 80, 80)
    End With
        
    'Antenna Label
    For p = 1 To SectorRowCount
        If InStr(Cells(p + 8, ArrFloor).Value, "L") And Cells(p + 8, ArrFloor).Value <> "LG" Then
            Cells(p + 8, 4).Interior.Color = RGB(196, 215, 155)
        Else
            Cells(p + 8, 4).Interior.Color = RGB(235, 241, 222)
        End If
    Next
    
    'Border
    Range("D9:AB" & 8 + SectorRowCount).Borders(xlInsideHorizontal).LineStyle = XlLineStyle.xlContinuous
    Range("D9:AB" & 8 + SectorRowCount).Borders(xlInsideVertical).LineStyle = XlLineStyle.xlContinuous
    
    excel.Application.DisplayAlerts = False
    MergeRow = 1
    For FRow = 1 To SectorRowCount
    
        If Cells(FRow + 8, 3).Value = Cells(FRow + 9, 3).Value Then
            If MergeState = 0 Then
                MergeState = 1
                MergeRow = FRow
           End If
        End If
        
        
        If Cells(FRow + 8, 3).Value <> Cells(FRow + 9, 3).Value Then
            Range(Cells(FRow + 8, 3), Cells(FRow + 8, 28)).Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlDouble
            
            If MergeState = 1 Then
                Range(Cells(MergeRow + 8, 3), Cells(FRow + 8, 3)).Merge
                MergeState = 0
            End If
        End If
    
    Next
    excel.Application.DisplayAlerts = True
    
    Range(Cells(9, 1), Cells(SectorRowCount + 8, 3)).HorizontalAlignment = xlCenter
    Range(Cells(9, 1), Cells(SectorRowCount + 8, 3)).VerticalAlignment = xlCenter

    Range("D9:AB" & 8 + SectorRowCount).BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
    
    With Range("AB9:AB" & 8 + SectorRowCount)
        .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
        .Borders.LineStyle = XlLineStyle.xlContinuous
        .Borders.Weight = xlMedium
        .HorizontalAlignment = xlCenter
    End With

    
End Sub
