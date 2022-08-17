Attribute VB_Name = "DC_AutoNumbering"
Option Explicit
Public Sub ActivateForm()
Attribute ActivateForm.VB_ProcData.VB_Invoke_Func = "W"

Dim vsoLayer As Visio.Layer

    ActiveWindow.DeselectAll
    
    AutoPage = ActivePage.name
    If ActivePage.Layers.Count > 0 Then
        For i = 1 To ActivePage.Layers.Count
            UserFormAutoNum.ListLayer.AddItem ActivePage.Layers(i).name
            Set vsoLayer = Application.ActiveWindow.Page.Layers.item(i)
            If vsoLayer.CellsC(visLayerLock).FormulaU = "1" Then
                UserFormAutoNum.ListLayer.Selected(i - 1) = True
            End If
        Next
    End If
    

    AutoNum = 0
    With UserFormAutoNum
        .Left = excel.Application.Left + excel.Application.Width - .Width
        '.Top = Excel.Application.Top
        '+ (0.5 * Excel.Application.Height) - (0.5 * .Height)
        .Show vbModeless
    End With
    AutoNum = 0
End Sub

Public Sub AutoShapeNum()
Dim PauseTime, Start, Finish, TotalTime
Dim PreShapeName As String

    PauseTime = 600    ' Set duration.
    Start = Timer   ' Set start time.
    
    Do While Timer < Start + PauseTime
        DoEvents
        If ActiveWindow.Selection.Count = 0 Then
            PreShapeName = ""
        End If
        If ActiveWindow.Selection.Count > 1 Then
            MsgBox "Plase do not select multiple shapes."
            End
        Else
            For Each vsoShape In ActiveWindow.Selection
                If Not IsNull(vsoShape.Cells("Prop.item_no").Formula) And _
                vsoShape.name <> PreShapeName Then
                    AutoNum = AutoNum + 1
                    UserFormAutoNum.TextBoxAntNum.Value = AutoNum
                    UserFormAutoNum.LabelNextNum.Caption = "Next Number :" & AutoNum + 1
                    vsoShape.Cells("Prop.item_no").Formula = AutoNum
                    PreShapeName = vsoShape.name
                End If
            Next
        End If
    Loop

End Sub

Sub PlusNum(AddUpNum As Integer)

    CheckError (ErrItemNumNotInteger)
    
    For Each vsoShape In ActiveWindow.Selection
        If (vsoShape.CellsSRC(visSectionProp, 1, visCustPropsValue)) > 0 Then
                vsoShape.Cells("Prop.item_no").Formula = vsoShape.Cells("Prop.item_no").Formula + AddUpNum
        End If
    Next
End Sub

    

    
