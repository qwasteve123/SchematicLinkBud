Attribute VB_Name = "C_VBAFunction"
Public Function CutStringPort(s As String) As String

    string1 = Right(s, Len(s) - InStr(1, s, "Connections.") - 11)
    CutStringPort = Left(string1, Len(string1) - 2)

End Function

Public Function CutShapeID(s As Variant)

    If IsNumeric(Right(s, Len(s) - InStr(1, s, "."))) Then
        CutShapeID = Right(s, Len(s) - InStr(1, s, "."))
    Else
        CutShapeID = 1
    End If
    
    
End Function

Public Function CouplerOrAnt(name As Variant)

If InStr(name, "Coupler") Or InStr(name, "Splitter") Then
    CouplerOrAnt = "C"
ElseIf InStr(name, "Ant") Then
    CouplerOrAnt = "A"
    'CouplerOrAnt = ""
ElseIf InStr(name, "connector") Then
    CouplerOrAnt = "S"
End If

    
End Function

Public Function ShapeCompType(name As Variant)

Select Case name
    Case "Connector"
        ShapeCompType = "S"
    Case "Omni Antenna", "Panel Antenna"
        'ShapeCompType = "A"
        ShapeCompType = ""
    Case "Coupler", "2 Way Splitter", "3 Way Splitter"
        ShapeCompType = "C"
End Select

End Function

Public Function ShapeCompFloor(floor As String)
    If InStr(floor, "L") And floor <> "LG" Then
        ShapeCompFloor = "L-" & Right(floor, Len(floor) - 1)
    Else
        ShapeCompFloor = floor
    End If
End Function

Public Function IsCouplePort(name As Variant)

If InStr(name, "direct_port") Then
    IsCouplePort = ">"
ElseIf InStr(name, "coupled_port") Then
    IsCouplePort = "^"
ElseIf InStr(name, "2way_output_port") Then
    IsCouplePort = "'"
ElseIf InStr(name, "3way_output_port") Then
    IsCouplePort = "*"
End If

End Function

Public Function calBudgetLoss(MaterialRow As Variant) As Double
    
Select Case FreqChoice
    Case "2690"
        calBudgetLoss = MaterialRow(ArrLCF12) * LossB26LCF12 + _
                        MaterialRow(ArrLCF78) * LossB26LCF78 + _
                        MaterialRow(ArrLCF114) * LossB26LCF114
    Case "3500"
        calBudgetLoss = MaterialRow(ArrLCF12) * LossB35LCF12 + _
                        MaterialRow(ArrLCF78) * LossB35LCF78 + _
                        MaterialRow(ArrLCF114) * LossB35LCF114
    Case Else
        calBudgetLoss = MaterialRow(ArrLCF12) * LossB26LCF12 + _
                        MaterialRow(ArrLCF78) * LossB26LCF78 + _
                        MaterialRow(ArrLCF114) * LossB26LCF114
                        
            
                        
End Select

    calBudgetLoss = calBudgetLoss + _
                    MaterialRow(ArrJumper) * LossJumper + _
                    MaterialRow(Arr2WaySplitter) * Loss2way + _
                    MaterialRow(Arr3WaySplitter) * Loss3way + _
                    MaterialRow(ArrC6Thr) * LossC6Thr + _
                    MaterialRow(ArrC6Couple) * LossC6Couple + _
                    MaterialRow(ArrC10Thr) * LossC10Thr + _
                    MaterialRow(ArrC10Couple) * LossC10Couple + _
                    MaterialRow(ArrC15Thr) * LossC15Thr + _
                    MaterialRow(ArrC15Couple) * LossC15Couple + _
                    MaterialRow(ArrC20Thr) * LossC20Thr + _
                    MaterialRow(ArrC20Couple) * LossC20Couple - _
                    MaterialRow(ArrAntGain) + _
                    MaterialRow(ArrCombiner) * LossCombiner + _
                    MaterialRow(ArrHybrid) * LossHybrid
                    
    calBudgetLoss = -calBudgetLoss
    
    If ShapeDataList(BudInRow, shdCompType) = "Panel Antenna" Then
        calBudgetLoss = calBudgetLoss - FSPL_lift + RSRP_output
    Else
        calBudgetLoss = calBudgetLoss - FSPL + RSRP_output
    End If
    
End Function

Public Function SetLabelValue(p As Integer)

Dim item As Variant
    
        Select Case ShapeDataList(p, shdCompType)
        
             Case "Connector"
                
                SetLabelValue = labvConnectors
        
            Case "Omni Antenna", "Panel Antenna"
            
                SetLabelValue = labvAntenna
                
            Case "Coupler", "2 Way Splitter", "3 Way Splitter"
            
                SetLabelValue = labvCoupler
                
        End Select
        
        For item = LBound(FloorList) To UBound(FloorList)
            If FloorList(item) = ShapeDataList(p, shdFloor) Then
                SetLabelValue = SetLabelValue + item * labvFloorList_W
            End If
        Next
        
        If Not IsNull(ShapeDataList(p, shdItemNo)) Then
            If ShapeDataList(p, shdItemNo) <> "" Then
                If IsNumeric(ShapeDataList(p, shdItemNo)) Then
                    SetLabelValue = SetLabelValue + ShapeDataList(p, shdItemNo)
                Else
                    SetLabelValue = SetLabelValue + 1
                End If
            Else
                SetLabelValue = SetLabelValue
            End If
        Else
            SetLabelValue = SetLabelValue
        End If
End Function

Function FloorValue(floor As Variant)

Dim i As Integer
Dim tempfloor As Integer

If IsNumeric(Left(floor, 1)) Then
    If IsNumeric(floor) Then
        
        FloorValue = FloorValue + labv1st_Floor + labvFloorW * floor
        
    Else
        For i = 2 To Len(floor)
            If Not IsNumeric(Left(floor, i)) Then
                tempfloor = Left(floor, i - 1)
                Exit For
            End If
        Next
        
        FloorValue = FloorValue + labv1st_Floor + labvFloorW * tempfloor + labvMiddleW
    
    End If
    
Else
    If InStr(floor, "G") Then
        If floor = "G" Then
            FloorValue = labvG_Floor
        ElseIf floor = "LG" Then
            FloorValue = labvG_Floor + 1 * labvFloorW
        ElseIf floor = "UG" Then
            FloorValue = labvG_Floor + 2 * labvFloorW
        Else
            FloorValue = labvG_Floor + 3 * labvFloorW
        End If
                
    ElseIf InStr(floor, "B") Then
        If IsNumeric(Right(floor, Len(floor) - 1)) Then
            FloorValue = FloorValue + labvG_Floor + labvBasementW * _
            Right(floor, Len(floor) - 1)
        Else
'            FloorValue = FloorValue + labvG_Floor + labvBasementW * _
'            Right(floor, Len(floor) - 1) + labvMiddleW
            
            FloorValue = FloorValue + labvG_Floor + labvBasementW
        End If
            
    ElseIf InStr(floor, "R") Then
        If floor = "R" Then
            FloorValue = FloorValue + labvRoof
        ElseIf floor = "UR" Then
            FloorValue = FloorValue + labvRoof + labvFloorW
            
        ElseIf floor = "MR" Then
            FloorValue = FloorValue + labvRoof + labvMiddleW
        Else
            FloorValue = FloorValue + labvRoof
        End If
            
    ElseIf InStr(floor, "L") Then
        If IsNumeric(Right(floor, Len(floor) - 1)) Then
            FloorValue = FloorValue + labvLift + labvLiftW * _
            Right(floor, Len(floor) - 1)
         Else
             FloorValue = FloorValue + labvLift
         End If
    End If
End If
             
End Function

Public Function SectorStartBlock(StartBlock As Variant)

Dim i As Integer

    For i = LBound(SectorList) To UBound(SectorList)
    
        If SectorList(i) = StartBlock Then
        
            SectorStartBlock = "Sector " & i + 1
            
        End If
        
    Next

End Function
