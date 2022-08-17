Attribute VB_Name = "DD_HighlightComponent"
Public Sub ActivateForm()
Attribute ActivateForm.VB_ProcData.VB_Invoke_Func = "Q"

     Erase ShapeDataList, FloorList


    For PageNum = 1 To ActiveDocument.Pages.Count

        Call ShapeData(PageNum)
    
    Next PageNum
    
    Call CountFloor
    
    

    With HighlightShapeForm
        .Combo_floor.List = FloorList
        .Combo_type.List = Array("2 Way Splitter", "3 Way Splitter", "Connector", "Coupler", "Omni Antenna", "Panel Antenna")
        .Combo_cable.List = Array("LCF4", "LCF5", "LCF6", "Jumper")
        .Left = excel.Application.Left + excel.Application.Width - .Width
        .Top = excel.Application.Top
        .Show vbModeless
    End With
    
    RelationMaxNo = 0
    ShapeMaxRow = 0
    FloorMaxRow = 0
   

End Sub

Public Sub HighLight(target_floor, target_comp_type As String, target_cable As String)
Dim floor, comp_type, cable_type As String
Dim error_circle_arr() As Shape
Dim find_comp_count As Integer
    
    For PageNum = 1 To ActiveDocument.Pages.Count
        
         Set vsoShapes = ActiveDocument.Pages.item(PageNum).Shapes
         
        For Each vsoShape In vsoShapes
            
            floor = CStr(vsoShape.CellsSRC(visSectionProp, 1, visCustPropsValue).Formula)
            If floor <> "" Then
                floor = Right(floor, Len(floor) - 1)
                floor = Left(floor, Len(floor) - 1)
            End If
            comp_type = CStr(vsoShape.CellsSRC(visSectionProp, 2, visCustPropsValue).Formula)
            
            If floor = target_floor Then
                If InStr(comp_type, target_comp_type) And vsoShape <> "" Then
                
                    find_comp_count = find_comp_count + 1
                
                    If InStr(comp_type, "Connector") Then
                        cable_type = CStr(vsoShape.CellsSRC(visSectionProp, 3, visCustPropsValue).ResultStr(visNone))
                        If InStr(cable_type, target_cable) Then
                            Call CircleItem(vsoShape, PageNum)
                            Call AddToList(vsoShape, PageNum, find_comp_count - 1)
                        Else
                            find_comp_count = find_comp_count - 1
                        End If
                    Else
                    
                        Call CircleItem(vsoShape, PageNum)
                        Call AddToList(vsoShape, PageNum, find_comp_count - 1)
                    End If
                End If
            End If
            
skiploop:
            
            floor = ""
            
            comp_type = ""
        Next
    
    Next PageNum

    PageNum = 0
    
    HighlightShapeForm.Label_CompCount.Caption = HighlightShapeForm.ListBox_Comp.ListCount

End Sub

Public Sub CircleItem(vsoShape As Shape, PageNum As Integer)

Dim CoorX As Double
Dim CoorY As Double
Dim comp_type As String

comp_type = CStr(vsoShape.CellsSRC(visSectionProp, 2, visCustPropsValue).Formula)


CoorX = vsoShape.Cells("PinX").Result("")
CoorY = vsoShape.Cells("PinY").Result("")

On Error GoTo skipprocedure
If InStr(comp_type, "Connector") Then
    CoorX = vsoShape.Cells("BeginX").Result("")
    CoorY = vsoShape.Cells("BeginY").Result("")
End If

skipprocedure:
ActiveDocument.Pages.item(PageNum).Drop ActiveDocument.Masters.ItemU("Error Circle"), CoorX, CoorY

End Sub

Public Sub AddToList(vsoShape As Shape, PageNum As Integer, list_index As Integer)

On Error GoTo skipaddlist:
With HighlightShapeForm.ListBox_Comp
    .AddItem CStr(vsoShape.CellsSRC(visSectionTextField, 0, 0).ResultStr(0))
    .List(list_index, 1) = vsoShape.name
    .List(list_index, 2) = PageNum & "_" & ActiveDocument.Pages.item(PageNum).name
    
End With

skipaddlist:


End Sub

