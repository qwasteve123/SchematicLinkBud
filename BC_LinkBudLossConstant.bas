Attribute VB_Name = "BC_LinkBudLossConstant"
'Output 43 dBm, 4 Operators, RSRP
Public Const Loss_RSRPB26 As Double = 6.208
Public Const Loss_RSRPB35 As Double = 2.844
'15m free space path loss, with fade margin 8.7 dBm, Power to n = 2.2
Public Const Loss_FSPLB26 As Double = -75.61
Public Const Loss_FSPLB35 As Double = -77.9
'Coaxial cable
Public Const LossB26LCF12 As Double = 0.124
Public Const LossB26LCF78 As Double = 0.0653
Public Const LossB26LCF114 As Double = 0.049

Public Const LossB35LCF12 As Double = 0.147
Public Const LossB35LCF78 As Double = 0.0795
Public Const LossB35LCF114 As Double = 0.058
'Jumper
Public Const LossJumper As Double = 0.5
'2 and 3 way splitter
Public Const Loss2way As Double = 3.6
Public Const Loss3way As Double = 5.6
'coupler
Public Const LossC6Thr As Double = 1.7
Public Const LossC6Couple As Double = 7

Public Const LossC10Thr As Double = 1
Public Const LossC10Couple As Double = 11.3

Public Const LossC15Thr As Double = 0.5
Public Const LossC15Couple As Double = 16.3

Public Const LossC20Thr As Double = 0.2
Public Const LossC20Couple As Double = 21.3

'Not used, have not added in the function
Public Const LossHybrid As Double = 3.1

Public Const LossCombiner As Double = 1

        
    
