Attribute VB_Name = "DB_CalSchematic"



Option Explicit
Public Sub GetSector()


    For i = 1 To ShapeMaxRow

        If InStr(ShapeDataList(i, shdCompName), "Start") Then
            If ShapeDataList(i, shdItemNo) = FreqChoice Or ShapeDataList(i, shdItemNo) = "No preference" Then
                    SectorMaxNum = SectorMaxNum + 1
                    ReDim Preserve SectorList(SectorMaxNum)
                    SectorList(SectorMaxNum - 1) = ShapeDataList(i, shdCompName)
            End If
        End If
        
    Next

End Sub
 Public Sub FindLinkage()

    LinkNumOfRow = 1

    For LinkRow = 1 + LinkSecRow To ShapeMaxRow

        If ShapeDataList(LinkRow, shddata3) <> 0 And ShapeDataList(LinkRow, shdCompType) = "Connector" Then

            LinkSecRow = LinkRow

            LinkageList(LinkNumOfRow, lltFirstName) = ShapeDataList(LinkRow, shdCompName)
            LinkageList(LinkNumOfRow, lltFirstNum) = LinkRow

            For LinkSecRow = LinkRow + 1 To ShapeMaxRow

                If ShapeDataList(LinkSecRow, shddata3) = ShapeDataList(LinkRow, shddata3) Then

                    TestLinkRow = LinkSecRow

                    LinkageList(LinkNumOfRow, lltSecondName) = ShapeDataList(LinkSecRow, shdCompName)
                    LinkageList(LinkNumOfRow, lltSecondNum) = LinkSecRow

                    LinkNumOfRow = LinkNumOfRow + 1
                    
                    CheckError (ErrLinkage)

                End If
            Next


        End If
    Next

    LinkNumOfRow = LinkNumOfRow - 1
    
    If LinkNumOfRow > 0 Then

        Call ConnectLinkage
        
    End If

 End Sub
 
 Public Sub ConnectLinkage()
 'Merge linkage
    
    RelationMaxNo = RelationMaxNo - 1
     For j = 1 To LinkNumOfRow
        For i = 1 To RelationMaxNo
            
            If Relation(i, relConnectors) = LinkageList(j, lltFirstName) Or _
            Relation(i, relConnectors) = LinkageList(j, lltSecondName) Then
                If InStr(Relation(i, relfromport), "cont") Then
                    LinkageToComp = Relation(i, reltocomp)
                    LinkageToPort = Relation(i, reltoport)
                    
                ElseIf InStr(Relation(i, reltoport), "cont") Then
                    LinkageFromComp = Relation(i, relfromcomp)
                    LinkageFromPort = Relation(i, relfromport)
                    LinkageConnectors = Relation(i, relConnectors)
                End If
                
                For k = 1 To relColMax
                    Relation(i, k) = "xxx"
                Next
            End If
            

        Next
        RelationMaxNo = RelationMaxNo + 1
        
        Relation(RelationMaxNo, relfromcomp) = LinkageFromComp
        Relation(RelationMaxNo, relfromport) = LinkageFromPort
        Relation(RelationMaxNo, relConnectors) = LinkageConnectors
        Relation(RelationMaxNo, reltocomp) = LinkageToComp
        Relation(RelationMaxNo, reltoport) = LinkageToPort
        
    Next
 End Sub
 
 Public Sub ShuffleRelation()

Dim i As Integer
Dim temp(MatListColMax) As Variant

    AntCount = 1
    
    For i = 1 To RelationMaxNo
        If InStr(1, Relation(i, relfromport), "start_port") > 0 Then
            
            SectorName = SectorStartBlock(Relation(i, relfromcomp))
            
            UpPath = "Start" & "-" & _
            "S." & CutShapeID(Relation(i, relConnectors)) & "-" & _
            CouplerOrAnt(Relation(i, reltocomp)) & "." & CutShapeID(Relation(i, reltocomp))
            
            
            LastPath = Relation(i, reltocomp)
            
                Call CountMatCoupler(i, LastPath, temp)
                Call CountMatConnector(i, LastPath, temp)
                 Call LinkBudget(temp, Relation(i, reltocomp))
            
            Call ShuffleRelation2(UpPath, LastPath, temp)
            
            Erase temp()
            
        End If
        
        SectorName = ""
        
    Next
          
End Sub

Public Sub ShuffleRelation2(UpPath As String, LastPath2 As Variant, step() As Variant)

Dim PathRoute As String
Dim j As Integer
Dim MaterialRow(MatListColMax) As Variant

If InStr(1, LastPath2, "Ant") > 0 Then

    PathRoute = UpPath
    
    LinkPath(AntCount, lkpAntShapeName) = LastPath2
    LinkPath(AntCount, lkpLinkPath) = UpPath
    
    For t = 1 To MatListColMax
        MaterialRow(t) = step(t)
    Next
    
    Call CountMatAnt(j, LastPath2, MaterialRow)
    
    For t = 1 To MatListColMax
        MaterialList(AntCount, t) = MaterialRow(t)
    Next
    
    AntCount = AntCount + 1
    
    Call LinkBudget(MaterialRow, LastPath2)

Else
    
    For j = 1 To RelationMaxNo - 1

            If Relation(j, 1) = LastPath2 Then
                    
                    PathRoute = UpPath
                    
                    PathRoute = PathRoute & IsCouplePort(Relation(j, relfromport)) & "-" & "S." & CutShapeID(Relation(j, relConnectors)) & _
                    "-" & CouplerOrAnt(Relation(j, reltocomp)) & "." & CutShapeID(Relation(j, reltocomp))
                    
                    For t = 1 To MatListColMax
                        MaterialRow(t) = step(t)
                    Next
                    
                            
                            Call CountMatCoupler(j, LastPath2, MaterialRow)
                            
                            Call CountMatConnector(j, LastPath2, MaterialRow)
                            
                            Call LinkBudget(MaterialRow, Relation(j, reltocomp))
                            
                   
                   
                    Call ShuffleRelation2(PathRoute, Relation(j, reltocomp), MaterialRow)
                           
            End If
    Next
    
End If

End Sub

Public Sub LinkBudget(MaterialRow As Variant, LastPathLink As Variant)


For BudInRow = 1 To ShapeMaxRow

    If ShapeDataList(BudInRow, shdCompName) = LastPathLink Then

        If ShapeDataList(BudInRow, shdStage) < 1 Or InStr(LastPathLink, "Ant") Then
           
            ShapeDataList(BudInRow, shdStage) = 1
            ShapeDataList(BudInRow, shdLinkBudget) = Format(calBudgetLoss(MaterialRow), "0.00")

        End If
        
    End If
    
Next

End Sub


Public Sub CountMatAnt(RelationRowNo As Integer, LastPath2 As Variant, MaterialRow() As Variant)


    For MatInRow = 1 To ShapeMaxRow

        If LastPath2 = ShapeDataList(MatInRow, shdCompName) Then

             MaterialRow(ArrAntGain) = ShapeDataList(MatInRow, shddata1) 'Ant Gain
             
             MaterialRow(ArrAntLabel) = ShapeDataList(MatInRow, shdCompLabel)
             
             MaterialRow(ArrFloor) = ShapeDataList(MatInRow, shdFloor)
             
             MaterialRow(ArrAntShapeName) = ShapeDataList(MatInRow, shdCompName)
             
             MaterialRow(ArrLabelIDValue) = ShapeDataList(MatInRow, shdLabelIDValue)
             
             MaterialRow(ArrSector) = SectorName
             
             BudInRow = MatInRow

        End If

    Next

End Sub

Public Sub CountMatConnector(RelationRowNo As Integer, LastPath2 As Variant, MaterialRow() As Variant)

    For MatInRow = 1 To ShapeMaxRow

        If Relation(RelationRowNo, relConnectors) = ShapeDataList(MatInRow, shdCompName) Then 'For Connectors


            Select Case ShapeDataList(MatInRow, shddata1)

                Case "Jumper"
                    MaterialRow(ArrJumper) = MaterialRow(ArrJumper) + 1
                Case "LCF4"
                    MaterialRow(ArrLCF12) = CInt(MaterialRow(ArrLCF12)) + ShapeDataList(MatInRow, shddata2)
                Case "LCF5"
                    MaterialRow(ArrLCF78) = CInt(MaterialRow(ArrLCF78)) + ShapeDataList(MatInRow, shddata2)
                    MaterialRow(ArrJumper) = MaterialRow(ArrJumper) + 2
                Case "LCF6"
                    MaterialRow(ArrLCF114) = CInt(MaterialRow(ArrLCF114)) + ShapeDataList(MatInRow, shddata2)
                    MaterialRow(ArrJumper) = MaterialRow(ArrJumper) + 2
            End Select

        End If

    Next
    
End Sub

Public Sub CountMatCoupler(RelationRowNo As Integer, LastPath2 As Variant, MaterialRow() As Variant)

    For MatInRow = 1 To ShapeMaxRow

        If Relation(RelationRowNo, relfromcomp) = ShapeDataList(MatInRow, shdCompName) Then 'For coupler or splitter

            Select Case ShapeDataList(MatInRow, shdCompType)

                Case "Coupler"

                    Select Case ShapeDataList(MatInRow, shddata1)

                        Case Is = 6 '________________________________________________

                            Select Case IsCouplePort(Relation(RelationRowNo, relfromport))
                                Case ">" 'Through Port
                                    MaterialRow(ArrC6Thr) = MaterialRow(ArrC6Thr) + 1
                                Case "^" 'Couple Port
                                    MaterialRow(ArrC6Couple) = MaterialRow(ArrC6Couple) + 1
                            End Select

                        Case Is = 10 '________________________________________________

                            Select Case IsCouplePort(Relation(RelationRowNo, relfromport))
                                Case ">" 'Through Port
                                    MaterialRow(ArrC10Thr) = MaterialRow(ArrC10Thr) + 1
                                Case "^" 'Couple Port
                                    MaterialRow(ArrC10Couple) = MaterialRow(ArrC10Couple) + 1
                            End Select

                        Case Is = 15 '________________________________________________

                            Select Case IsCouplePort(Relation(RelationRowNo, relfromport))
                                Case ">" 'Through Port
                                    MaterialRow(ArrC15Thr) = MaterialRow(ArrC15Thr) + 1
                                Case "^" 'Couple Port
                                    MaterialRow(ArrC15Couple) = MaterialRow(ArrC15Couple) + 1
                            End Select

                        Case Is = 20 '________________________________________________

                            Select Case IsCouplePort(Relation(RelationRowNo, relfromport))
                                Case ">" 'Through Port
                                    MaterialRow(ArrC20Thr) = MaterialRow(ArrC20Thr) + 1
                                Case "^" 'Couple Port
                                    MaterialRow(ArrC20Couple) = MaterialRow(ArrC20Couple) + 1
                            End Select

                    End Select

                Case "2 Way Splitter"

                    MaterialRow(Arr2WaySplitter) = MaterialRow(Arr2WaySplitter) + 1

                Case "3 Way Splitter"

                    MaterialRow(Arr3WaySplitter) = MaterialRow(Arr3WaySplitter) + 1
                    
                Case "Combiner"
                
                    MaterialRow(ArrCombiner) = MaterialRow(ArrCombiner) + 1
                    
                Case "Hybrid"
                
                    MaterialRow(ArrHybrid) = MaterialRow(ArrHybrid) + 1

            End Select

        End If

    Next
    
End Sub

Sub LabelToLinkBud()


Dim vsoLinkBud As Visio.Cell
Dim NamingStart As Integer

'PageName = ActiveDocument.Pages.item(PageNum).Index & "_"
PageName = ActivePage.Index & "_"

'Set vsoShapes = ActiveDocument.Pages.item(PageNum).Shapes

Set vsoShapes = ActivePage.Shapes

    For i = 1 To ShapeMaxRow

'        If ShapeDataList(i, shdLinkBudget) <> 0 And _
'        ShapeDataList(i, shdPageNum) = ActiveDocument.Pages.item(PageNum).name Then
        
        
        If ShapeDataList(i, shdLinkBudget) <> 0 And _
        ShapeDataList(i, shdPageNum) = ActivePage.name Then

            For Each vsoShape In vsoShapes

                If PageName & vsoShape.name = ShapeDataList(i, shdCompName) Then

                    If vsoShape.Cells("User.link_bud_text").Formula = 0 And NamingState = 0 Then
                        NamingStart = 1
                    End If
                    
                    If NamingStart = 1 Then
                        vsoShape.Cells("User.link_bud_text").Formula = ShapeDataList(i, shdLinkBudget)
                    Else
                        vsoShape.Cells("User.link_bud_text").Formula = 0

                    End If
                    
                    If FSPL = 0 Then
                        vsoShape.Cells("User.link_bud_unit").Formula = """dB"""
                    Else
                        vsoShape.Cells("User.link_bud_unit").Formula = """dBm"""
                    End If

                End If

            Next

        End If

    Next

    
    

End Sub


