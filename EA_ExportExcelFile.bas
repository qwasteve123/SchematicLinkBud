Public Sub ExcelFile()
Dim fd As FileDialog
Dim file_path As Variant
Dim wb_name As String

Set fd = excel.Application.FileDialog(msoFileDialogFolderPicker)
fd.InitialFileName = Environ("Userprofile") & "\Desktop"
fd.AllowMultiSelect = False
Actionclicked = fd.Show
If Actionclicked Then
    file_path = fd.SelectedItems(1)
Else
    MsgBox "Please choose a folder directory to export your link budget excel file."
    GoTo exitsub
End If

Workbooks.Add.SaveAs fileName:=file_path & "\" & "link_buget.xlsx"
'Environ ("userprofile") & "\Desktop\trial4.xlsx", ConflictResolution:=xlLocalSessionChanges

Call ExcelFileInfo
Call FormalTemplate
Call BOMTemplate

Range("A1") = "hello"

Workbooks("link_buget.xlsx").Save
Workbooks("link_buget.xlsx").Close

exitsub:
End Sub

Sub ExcelFileInfo()


'_____________________________________________________________________________________
ActiveSheet.name = "Relation"

Cells(1, relfromcomp).Value = "From"
Cells(1, relfromport).Value = "From port"
Cells(1, relConnectors).Value = "Connectors"
Cells(1, reltocomp).Value = "To"
Cells(1, reltoport).Value = "To port"

Range(Cells(2, 1), Cells(RelationMaxNo + 1, relColMax)).Clear
Range(Cells(2, 1), Cells(RelationMaxNo + 1, relColMax)) = Relation

Range(Cells(1, 1), Cells(RelationMaxNo + 1, relColMax)).Columns.AutoFit

'_____________________________________________________________________________________

ActiveWorkbook.Sheets.Add(After:=Worksheets(ActiveWorkbook.Worksheets.Count)).name = "Link Path"

Cells(1, 1).Value = "Ant Name"
Cells(1, 2).Value = "Link Path"

Range(Cells(2, 1), Cells(AntCount + 1, lkpColMax)).Clear
Range(Cells(2, 1), Cells(AntCount + 1, lkpColMax)) = LinkPath

Range(Cells(1, 1), Cells(AntCount + 1, lkpColMax)).Columns.AutoFit

'_____________________________________________________________________________________


ActiveWorkbook.Sheets.Add(After:=Worksheets(ActiveWorkbook.Worksheets.Count)).name = "Shape Data"

Cells(1, shdLabelIDValue).Value = "Shape Label"
Cells(1, shdCompName).Value = "Shape Name"
Cells(1, shdCompLabel).Value = "Item Label"
Cells(1, shdItemNo).Value = "Item No."
Cells(1, shdFloor).Value = "Floor"
Cells(1, shdCompType).Value = "Component Type"
Cells(1, shddata1).Value = "Data #1"
Cells(1, shddata2).Value = "Data #2"
Cells(1, shddata3).Value = "Data #3"
Cells(1, shdStage).Value = "Label Exist"
Cells(1, shdLinkBudget).Value = "Link Budget"
Cells(1, shdPageNum).Value = "Page Name"

Range(Cells(2, 1), Cells(ShapeMaxRow + 1, shdColMax)).Clear
Range(Cells(2, 1), Cells(ShapeMaxRow + 1, shdColMax)) = ShapeDataList

Range(Cells(1, 1), Cells(ShapeMaxRow + 1, shdColMax)).Columns.AutoFit


For i = 1 To ShapeMaxRow
    If Cells(i + 1, shdCompType) = "Omni Antenna" Then
        If Cells(i + 1, shdItemNo) Mod 10 = 0 Then
            Cells(i + 1, shdCompLabel).NumberFormat = "0.00"
            Cells(i + 1, shdCompLabel).Value = Cells(i + 1, shdFloor) & "." & Cells(i + 1, shdItemNo)
        End If
    End If
    Cells(i + 1, shdCompLabel).HorizontalAlignment = xlHAlignRight
Next


With ActiveSheet.Sort

    .SortFields.Add Key:=Range("A1"), Order:=xlAscending
    .SetRange Range(Cells(1, 1), Cells(ShapeMaxRow + 1, shdColMax))
    .Header = xlYes
    .Apply

End With


'_____________________________________________________________________________________

ActiveWorkbook.Sheets.Add(After:=Worksheets(ActiveWorkbook.Worksheets.Count)).name = "Material List"

Cells(1, ArrLabelIDValue).Value = "Label ID"
Cells(1, ArrFloor).Value = "Floor"
Cells(1, ArrAntShapeName).Value = "Antenna Shape Name"
Cells(1, ArrAntLabel).Value = "Antenna Label"
Cells(1, ArrLCF12).Value = "LCF 12"
Cells(1, ArrLCF78).Value = "LCF 78"
Cells(1, ArrLCF114).Value = "LCF 114"
Cells(1, ArrJumper).Value = "Jumper"
Cells(1, Arr2WaySplitter).Value = "2-way Splitter"
Cells(1, Arr3WaySplitter).Value = "3-way Splitter"
Cells(1, ArrC6Thr).Value = "6dB Thr."
Cells(1, ArrC6Couple).Value = "6dB Couple"
Cells(1, ArrC10Thr).Value = "10dB Thr."
Cells(1, ArrC10Couple).Value = "10dB Couple"
Cells(1, ArrC15Thr).Value = "15dB Thr."
Cells(1, ArrC15Couple).Value = "15dB Couple"
Cells(1, ArrC20Thr).Value = "20dB Thr."
Cells(1, ArrC20Couple).Value = "20dB Couple"
Cells(1, ArrAntGain).Value = "Ant Gain"
Cells(1, ArrHybrid).Value = "Hybrid"
Cells(1, ArrCombiner).Value = "Combiner"
Cells(1, ArrSector).Value = "Sector"

Range(Cells(2, 1), Cells(AntCount, MatListColMax)).Clear
Range(Cells(2, 1), Cells(AntCount, MatListColMax)) = MaterialList

Range(Cells(1, 1), Cells(AntCount, MatListColMax)).Columns.AutoFit

For i = 1 To AntCount
    If Right(CStr(MaterialList(i, ArrAntLabel)), 1) = "0" And Left(MaterialList(i, ArrAntLabel), 1) <> "L" Then
        Cells(i + 1, ArrAntLabel).NumberFormat = "0.00"
    End I
    Cells(i + 1, ArrAntLabel).HorizontalAlignment = xlHAlignRight
Next

With ActiveSheet.Sort

    .SortFields.Add Key:=Range("A1"), Order:=xlAscending
    .SetRange Range("A1:V" & 1 + AntCount)
    .Header = xlYes
    .Apply

End With




'_____________________________________________________________________________________


End Sub
