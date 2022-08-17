Attribute VB_Name = "DA_GetData"
Option Explicit
Public Sub GetInformation() 'In DataShape Sub Loop

Dim page_num As Integer
Dim vsoShape As Visio.Shape
Dim vsoShapes As Visio.Shapes
    
For page_num = 1 To ActiveDocument.Pages.Count
    Set vsoShapes = ActiveDocument.Pages.item(page_num).Shapes
    For Each vsoShape In vsoShapes
        If InStr(vsoShape.name, "ChoiceBlock") Then
            FSPL = Format(vsoShape.Cells("User.FSPL").Result(""), "0.000")
            FSPL_lift = Format(vsoShape.Cells("User.FSPL_Lift").Result(""), "0.000")
            RSRP_output = Format(vsoShape.Cells("User.RSRP_output").Result(""), "0.000")
            FreqChoice = vsoShape.Cells("Prop.fspl_freq").Result("")
        End If
    Next
Next
End Sub



Public Sub ShapeConnection(ByVal PageNum As Integer)
    
            Set vsoShapes = ActiveDocument.Pages.item(PageNum).Shapes
            
            If RelationMaxNo > 1 Then
            
                RelationNo = RelationMaxNo
                
            Else
            
                RelationNo = 1
                
            End If
            
            PageName = ActiveDocument.Pages.item(PageNum).Index & "_"
            
        
            For vsoShapeNum = 1 To vsoShapes.Count
            
                Set vsoShape = vsoShapes(vsoShapeNum)
            
                Set vsoConnects = vsoShape.Connects
                
                If vsoConnects.Count > 0 Then
                    For i = 1 To vsoConnects.Count

                        Set vsoConnect = vsoConnects(i)
                        Set vsoConnectfromCell = vsoConnect.FromCell
                        Set vsoConnectToCell = vsoConnect.ToCell
                        
                        If vsoConnect.ToSheet.CellExists("prop.freq_choice", 0) = -1 Then
                            If vsoConnect.ToSheet.Cells("prop.freq_choice").Result(visNone) <> FreqChoice And _
                             vsoConnect.ToSheet.Cells("prop.freq_choice").ResultStr(visNone) <> "No preference" Then
                                GoTo skiploop
                            End If
                        End If
                        
                        Call CheckError(ErrConnector)
                        Call CheckError(ErrWalkGlue)
                        
                        
                        ConnectString = Left(vsoConnectfromCell.Formula, InStr(1, vsoConnectfromCell.Formula, ",") - 1)
                        ConnectString = Right(ConnectString, Len(ConnectString) - 8)

                        
        
                        If InStr(1, ConnectString, "cont") Then
                            If ConnectStringStatus = 0 Then
                                ConnectStringTemp = ConnectString
                            ElseIf ConnectStringStatus = 1 Then
                                ConnectToString = ConnectString
                            ElseIf ConnectStringStatus = 2 Then
                                ConnectFromString = ConnectString
                            End If

                        Else
                        

 
                                If InStr(1, ConnectString, "input") Or InStr(1, ConnectString, "ant") Then
                                
                                    ConnectToString = ConnectString
                                    ConnectStringStatus = 2
                                    
                                Else
                                
                                    ConnectFromString = ConnectString
                                    ConnectStringStatus = 1
                                    
                                End If
                                
                                If ConnectStringTemp <> "" Then
                                    If ConnectStringStatus = 1 Then
                                        ConnectToString = ConnectStringTemp
                                    ElseIf ConnectStringStatus = 2 Then
                                        ConnectFromString = ConnectStringTemp
                                    End If
                                End If
                                
                        End If

                    Next i
                        
                    ConnectFromPort = CutStringPort(ConnectFromString)
                    ConnectToPort = CutStringPort(ConnectToString)
                    
                    Call CheckError(ErrSamePort)
                    
                    ConnectFromName = PageName & Left(ConnectFromString, InStr(1, ConnectFromString, "!") - 1)
                    ConnectToName = PageName & Left(ConnectToString, InStr(1, ConnectToString, "!") - 1)
                    
                    ConnectFromID = CutShapeID(ConnectFromName)
                    ConnectToID = CutShapeID(ConnectToName)
                    
                    Relation(RelationNo, relfromcomp) = ConnectFromName
                    Relation(RelationNo, relfromport) = ConnectFromPort
                    
                    Relation(RelationNo, relConnectors) = PageName & vsoShape.name
                    
                    Relation(RelationNo, reltocomp) = ConnectToName
                    Relation(RelationNo, reltoport) = ConnectToPort
                    
                    RelationNo = RelationNo + 1
                    
                    ConnectFromName = ""
                    ConnectFromPort = ""
                    
                    ConnectToName = ""
                    ConnectToPort = ""
                    
                    ConnectStringStatus = 0
                    ConnectStringTemp = ""
                    
                    

                         
                End If
                
skiploop:
            Next vsoShapeNum
            
            i = 0
            
            RelationMaxNo = RelationNo

End Sub

Public Sub ShapeData(ByVal PageNum As Integer)

 Dim vsoShapes As Visio.Shapes
 Dim vsoShape As Visio.Shape
 Dim shdRowIndex As Integer
 Dim intCounter As Integer
 Dim vsoCell As Visio.Cell
 
    
 
        Set vsoShapes = ActiveDocument.Pages.item(PageNum).Shapes
        
        PageName = ActiveDocument.Pages.item(PageNum).Index & "_"
          
        For PageShapeIndex = 1 To vsoShapes.Count
        
            shdRowIndex = PageShapeIndex + ShapeMaxRow
         
            Set vsoShape = vsoShapes(PageShapeIndex)
             
             intRows = vsoShape.RowCount(visSectionProp)
             
             ShapeDataList(shdRowIndex, shdCompName) = PageName & vsoShape.name
             
             For intCounter = 0 To intRows - 1
             
                Set vsoCellValue = vsoShape.CellsSRC(visSectionProp, intCounter, visCustPropsValue)
                ShapeDataList(shdRowIndex, 4 + intCounter) = vsoCellValue.ResultStr(visNone)
                
            Next intCounter
                
                If IsNumeric(ShapeDataList(shdRowIndex, shdFloor)) Then
                    ShapeDataList(shdRowIndex, shdFloor) = CInt(ShapeDataList(shdRowIndex, shdFloor))
                End If
                
                ShapeDataList(shdRowIndex, shdCompLabel) = ShapeCompType(ShapeDataList(shdRowIndex, shdCompType)) & _
                                                                ShapeCompFloor(Format(ShapeDataList(shdRowIndex, shdFloor), 0)) & "." & _
                                                                Format(ShapeDataList(shdRowIndex, shdItemNo), 0)
                
                ShapeDataList(shdRowIndex, shdPageNum) = ActiveDocument.Pages.item(PageNum).name
                
                
            'Reset shdStage
            ShapeDataList(shdRowIndex, shdStage) = 0
            ShapeDataList(shdRowIndex, shdLinkBudget) = 0
            
            
            
           ' Call GetInformation(vsoShape)
            
'            If ShapeDataList(shdRowIndex, shdCompName) = "Connector" Then
'                PortRowCount = PortRowCount + 1
'                SamePort(PortRowCount, spPort1CoorX) = vsoshape.Cells("BeginX").Result("")
'                SamePort(PortRowCount, spPort1CoorX) = vsoshape.Cells("BeginY").Result("")
'
'                SamePort(PortRowCount, spPort2CoorX) = vsoshape.Cells("EndX").Result("")
'                SamePort(PortRowCount, spPort2CoorX) = vsoshape.Cells("EndY").Result("")
'            End If
         
         Next PageShapeIndex
         
        ShapeMaxRow = ShapeMaxRow + vsoShapes.Count
         
         
 
End Sub

Public Sub CountFloor()

Dim p As Integer
Dim FloorDup As Boolean

    For p = 1 To ShapeMaxRow
    
        FloorDup = False
            
        If FloorMaxRow = 0 Then
        
            FloorMaxRow = 1
            
            ReDim Preserve FloorList(FloorMaxRow)
            
        Else
                
            For FloorRow = 0 To FloorMaxRow
    
                If FloorList(FloorRow) = ShapeDataList(p, shdFloor) Or IsNull(ShapeDataList(p, shdFloor)) Then
                    FloorDup = True
                    Exit For
                End If
            Next
            
            If FloorDup = False Then
            
                FloorList(FloorMaxRow - 1) = ShapeDataList(p, shdFloor)
            
                FloorMaxRow = FloorMaxRow + 1
                
                ReDim Preserve FloorList(FloorMaxRow)
                
            End If
            
        End If

    Next
    If FloorMaxRow > 1 Then
        FloorMaxRow = FloorMaxRow - 2
    End If
    ReDim Preserve FloorList(FloorMaxRow)
    
    Call SortFloor
    
        
End Sub

Public Sub SortFloor()

Dim i As Integer
Dim j As Integer
Dim temp As Variant
    
For i = LBound(FloorList) To UBound(FloorList)
    For j = i + 1 To UBound(FloorList)
    
        If FloorValue(FloorList(i)) > FloorValue(FloorList(j)) Then
    
            temp = FloorList(i)
            FloorList(i) = FloorList(j)
            FloorList(j) = temp
        End If
    Next
Next

End Sub

Public Sub SetLabelID()

Dim LabelRow As Integer

    For LabelRow = 1 To ShapeMaxRow
        ShapeDataList(LabelRow, shdLabelIDValue) = SetLabelValue(LabelRow)
    Next

End Sub


    
