Attribute VB_Name = "DE_ChangeLabelFormat"
Option Explicit

Dim label_lift_common, label_normal_common As String
Dim label_lift_2_way, label_normal_2_way As String
Dim label_lift_3_way, label_normal_3_way As String
Dim label_lift_panel_ant, label_normal_panel_ant As String
Dim label_lift_coupler, label_normal_coupler As String
Dim label_lift_connector_cable, label_normal_connector_cable As String
Dim label_lift_connector_whole, label_normal_connector_whole As String

Sub temp()

    Call ChangeLabelName(False)

End Sub

Sub ChangeLabelName(label_is_lift As Boolean)

Dim vsoShape As Visio.Shape
Dim label_formula As String
Dim component_type

    
    
    Call GetLabel

    For Each vsoShape In ActiveWindow.Selection
        On Error GoTo skiploop
        
        label_formula = vsoShape.CellsSRC(visSectionUser, 0, 0).Formula
        
        Select Case vsoShape.CellsSRC(visSectionProp, 2, visCustPropsValue).ResultStr(visNone)
            Case "2 Way Splitter"
                If label_is_lift Then
                    label_formula = label_lift_2_way
                Else
                    label_formula = label_normal_2_way
                End If
            Case "3 Way Splitter"
                If label_is_lift Then
                    label_formula = label_lift_3_way
                Else
                    label_formula = label_normal_3_way
                End If
            Case "Coupler"
                If label_is_lift Then
                    label_formula = label_lift_coupler
                Else
                    label_formula = label_normal_coupler
                End If
            Case "Panel Antenna"
                If label_is_lift Then
                    label_formula = label_lift_panel_ant
                Else
                    label_formula = label_normal_panel_ant
                End If
            Case "Connector"
                If label_is_lift Then
                    label_formula = label_lift_connector_whole
                Else
                    label_formula = label_normal_connector_whole
                End If
        End Select

        vsoShape.CellsSRC(visSectionUser, 0, 0).Formula = label_formula
        label_formula = ""
        
skiploop:
    Next


End Sub

Sub GetLabel()

    label_lift_common = Chr(38) & "Right(Prop.floor, Len(Prop.floor) - 1)" & Chr(38) & Chr(34) & Chr(46) & Chr(34) & Chr(38) & "Prop.item_no & CHAR(10)" & Chr(38)
    label_normal_common = Chr(38) & "Prop.floor" & Chr(38) & Chr(34) & Chr(46) & Chr(34) & Chr(38) & "Prop.item_no & CHAR(10)" & Chr(38)


    label_lift_2_way = """L-C""" & label_lift_common & """3dB"""
    label_normal_2_way = """C""" & label_normal_common & """3dB"""
    
    label_lift_3_way = """L - C""" & label_lift_common & """5dB"""
    label_normal_3_way = """C""" & label_normal_common & """5dB"""
    
    label_lift_coupler = """L - C""" & label_lift_common & "Prop.coupling_loss" & Chr(38) & """dB"""
    label_normal_coupler = """C""" & label_normal_common & "Prop.coupling_loss" & Chr(38) & """dB"""
    
    label_lift_panel_ant = """L-""" & Chr(38) & "Right(Prop.floor, Len(Prop.floor) - 1)" & Chr(38) & Chr(34) & Chr(46) & Chr(34) & Chr(38) & "Prop.item_no"
    label_normal_panel_ant = "Prop.floor" & Chr(38) & Chr(34) & Chr(46) & Chr(34) & Chr(38) & "Prop.item_no"
    
    label_lift_connector_cable = """L - S""" & label_lift_common & "Prop.feeder_type& CHAR(32) &Prop.feeder_length" & Chr(38) & """m"""
    label_normal_connector_cable = """S""" & label_normal_common & "Prop.feeder_type& CHAR(32) &Prop.feeder_length" & Chr(38) & """m"""
    
    label_lift_connector_whole = "IF(User.feeder_index=0," & """J""" & Chr(44) & label_lift_connector_cable & ")"
    label_normal_connector_whole = "IF(User.feeder_index=0," & """J""" & Chr(44) & label_normal_connector_cable & ")"

End Sub
