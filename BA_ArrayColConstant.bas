
Public Enum ArrMatList

    ArrLabelIDValue = 1
    ArrAntShapeName = 2
    ArrFloor = 3
    ArrAntLabel = 4
    ArrLCF12 = 5
    ArrLCF78 = 6
    ArrLCF114 = 7
    ArrJumper = 8
    Arr2WaySplitter = 9
    Arr3WaySplitter = 10
    ArrC6Thr = 11
    ArrC6Couple = 12
    ArrC10Thr = 13
    ArrC10Couple = 14
    ArrC15Thr = 15
    ArrC15Couple = 16
    ArrC20Thr = 17
    ArrC20Couple = 18
    ArrHybrid = 19
    ArrCombiner = 20
    ArrAntGain = 21
    ArrSector = 22
    
    MatListColMax = 22

End Enum

Public Enum ArrShapeData

    shdLabelIDValue = 1
    shdCompName = 2
    shdCompLabel = 3
    shdItemNo = 4
    shdFloor = 5
    shdCompType = 6
    shddata1 = 7
    shddata2 = 8
    shddata3 = 9
    shdStage = 10
    shdLinkBudget = 11
    shdPageNum = 12
    shdShapeUID = 13
    
    
    shdColMax = 13

End Enum

Public Enum ArrRelation

    relfromcomp = 1
    relfromport = 2
    relConnectors = 3
    reltocomp = 4
    reltoport = 5
    
    relColMax = 5


End Enum


Public Enum ArrLinkPath

    lkpAntShapeName = 1
    lkpLinkPath = 2

    lkpColMax = 2

End Enum

Public Enum ArrLinkageList

    lltFirstNum = 1
    lltFirstName = 2
    lltSecondNum = 3
    lltSecondName = 4
    
    lltColMax = 4
    
End Enum

Public Enum LabelValue

    labvAntenna = 1000000
    labvCoupler = 2000000
    labvConnectors = 3000000
    labvFloorList_W = 1000
    
    labvG_Floor = 10000
    labv1st_Floor = 20000
    labvLift = 90000
    labvRoof = 50000
    labvBasementW = -1000
    labvFloorW = 1000
    labvLiftW = 1000
    labvMiddleW = 500
    
    
End Enum

Enum BondMaterial

    bomFloor = 1
    bomLCF12 = 2
    bomLCF78 = 3
    bomLCF114 = 4
    bomJumper = 5
    bom2WaySplitter = 6
    bom3WaySplitter = 7
    bomC6 = 8
    bomC10 = 9
    bomC15 = 10
    bomC20 = 11
    bomConnector12 = 12
    bomConnector78 = 13
    bomConnector114 = 14
    bomHybrid = 15
    bomCombiner = 16
    bomOmniAnt = 17
    bompanelAnt = 18
    
    BondMatColMax = 18
    
End Enum

Enum SamePortArr

    spPort1CoorX = 1
    spPort1CoorY = 2
    spPort2CoorX = 3
    spPort2CoorY = 4
    
End Enum
