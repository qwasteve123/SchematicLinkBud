Attribute VB_Name = "EC_ExportExcelFile3"
Option Explicit
Sub BOMTemplate()

    ActiveWorkbook.Sheets.Add(After:=Worksheets(ActiveWorkbook.Worksheets.Count)).name = "BOM_Floor"
    
    ActiveSheet.Cells.Clear

    Call BondOfMaterial
    Call TitleBar
    Call SumEquation
    Call BorderNColor
    
End Sub
Sub BorderNColor()
    
    'Title Bar
    Range(Cells(1, 1), Cells(1, BondMatColMax)).Interior.Color = RGB(206, 206, 206)
    'Border
    Range(Cells(1, 1), Cells(FloorMaxRow + 3, BondMatColMax)).BorderAround Weight:=xlMedium
    Range(Cells(1, 1), Cells(1, BondMatColMax)).BorderAround Weight:=xlMedium
    Range(Cells(1, 1), Cells(FloorMaxRow + 3, 1)).BorderAround Weight:=xlMedium
    
    With Range(Cells(FloorMaxRow + 3, 1), Cells(FloorMaxRow + 3, BondMatColMax)).Borders(xlEdgeTop)
        .Weight = xlMedium
        .LineStyle = XlLineStyle.xlDouble
    End With
    Range(Cells(1, 1), Cells(FloorMaxRow + 3, 1)).HorizontalAlignment = xlCenter

End Sub

Sub TitleBar()
    
    Cells(1, bomFloor).Value = "Floor"
    Cells(1, bomLCF12).Value = "LCF4"
    Cells(1, bomLCF78).Value = "LCF5"
    Cells(1, bomLCF114).Value = "LCF6"
    Cells(1, bomJumper).Value = "Jumper"
    Cells(1, bom2WaySplitter).Value = "2 Way Splitter"
    Cells(1, bom3WaySplitter).Value = "3 Way Splitter"
    Cells(1, bomC6).Value = "6 dB"
    Cells(1, bomC10).Value = "10 dB"
    Cells(1, bomC15).Value = "15 dB"
    Cells(1, bomC20).Value = "20 dB"
    Cells(1, bomConnector12).Value = "LCF4 Connectors"
    Cells(1, bomConnector78).Value = "LCF5 Connectors"
    Cells(1, bomConnector114).Value = "LCF6 Connectors"
    Cells(1, bomHybrid).Value = "Hybrid"
    Cells(1, bomCombiner).Value = "Combiner"
    Cells(1, bomOmniAnt).Value = "Omni Antenna"
    Cells(1, bompanelAnt).Value = "Panel Antenna"
    
End Sub

Sub SumEquation()
    
    Range(Cells(2, 1), Cells(FloorMaxRow + 2, BondMatColMax)) = BondMatList
    Range(Cells(1, 1), Cells(FloorMaxRow + 2, BondMatColMax)).Columns.AutoFit
    Cells(FloorMaxRow + 3, 1).Value = "Total"
    Range("B" & FloorMaxRow + 3).FormulaR1C1 = "=sum(R[" & -FloorMaxRow - 1 & "]C[0]:R[-1]C[0])"
    
    Range("B" & FloorMaxRow + 3).AutoFill Destination:= _
    Range("B" & FloorMaxRow + 3 & ":R" & FloorMaxRow + 3), Type:=xlFillCopy
End Sub


Public Sub BondOfMaterial()

    ReDim BondMatList(0 To FloorMaxRow, 1 To BondMatColMax)
    
    
    
    
    For j = 1 To ShapeMaxRow
    
        For i = 0 To FloorMaxRow
    
            
            If ShapeDataList(j, shdFloor) = FloorList(i) Then
            
                Select Case ShapeDataList(j, shdCompType)
                
 '__________________________________________________________________________________________________________________________
                'Coaxial cable
                    Case "Connector"

                        Select Case ShapeDataList(j, shddata1)
                        
                            Case "LCF4"
                            
                                BondMatList(i, bomLCF12) = BondMatList(i, bomLCF12) + ShapeDataList(j, shddata2) * SisoOrMimo
                                BondMatList(i, bomConnector12) = BondMatList(i, bomConnector12) + 2 * SisoOrMimo
                                
                            Case "LCF5"
                            
                                BondMatList(i, bomLCF78) = BondMatList(i, bomLCF78) + ShapeDataList(j, shddata2) * SisoOrMimo
                                BondMatList(i, bomConnector78) = BondMatList(i, bomConnector78) + 2 * SisoOrMimo
                                BondMatList(i, bomJumper) = BondMatList(i, bomJumper) + 2 * SisoOrMimo
                        
                            Case "LCF6"
                            
                                BondMatList(i, bomLCF114) = BondMatList(i, bomLCF114) + ShapeDataList(j, shddata2) * SisoOrMimo
                                BondMatList(i, bomConnector114) = BondMatList(i, bomConnector114) + 2 * SisoOrMimo
                                BondMatList(i, bomJumper) = BondMatList(i, bomJumper) + 2 * SisoOrMimo
                                
                            Case "Jumper"

                                BondMatList(i, bomJumper) = BondMatList(i, bomJumper) + SisoOrMimo
                                
                        End Select
'__________________________________________________________________________________________________________________________
                                
                                
                    Case "2 Way Splitter"
                    
                        BondMatList(i, bom2WaySplitter) = BondMatList(i, bom2WaySplitter) + 1 * SisoOrMimo
                        
                    Case "3 Way Splitter"
                    
                        BondMatList(i, bom3WaySplitter) = BondMatList(i, bom3WaySplitter) + 1 * SisoOrMimo
                        
                    Case "Coupler"
                    
                        Select Case ShapeDataList(j, shddata1)
                        
                            Case "6"
                            
                                BondMatList(i, bomC6) = BondMatList(i, bomC6) + 1 * SisoOrMimo
                        
                            Case "10"
                            
                                BondMatList(i, bomC10) = BondMatList(i, bomC10) + 1 * SisoOrMimo
                        
                            Case "15"
                            
                                BondMatList(i, bomC15) = BondMatList(i, bomC15) + 1 * SisoOrMimo
                                
                            Case "20"
                            
                                BondMatList(i, bomC20) = BondMatList(i, bomC20) + 1 * SisoOrMimo
                                
                         End Select
                        
'__________________________________________________________________________________________________________________________
                
                    Case "Hybrid"
                        
                        BondMatList(i, bomHybrid) = BondMatList(i, bomHybrid) + 1 * SisoOrMimo
                        
                    Case "Combiner"
                        
                        BondMatList(i, bomCombiner) = BondMatList(i, bomCombiner) + 1 * SisoOrMimo
                        
                'Omni Antenna
                    Case "Omni Antenna"
                        
                        BondMatList(i, bomOmniAnt) = BondMatList(i, bomOmniAnt) + 1 * SisoOrMimo
                        
                    Case "Panel Antenna"
                        
                        BondMatList(i, bompanelAnt) = BondMatList(i, bompanelAnt) + 1 * SisoOrMimo
                        Debug.Print ShapeDataList(j, shdCompName)
                        
                    End Select
                    
                End If
                
            Next
              
        Next
        
      For i = 0 To FloorMaxRow
      
        BondMatList(i, bomFloor) = FloorList(i)
        
    Next
      
End Sub
