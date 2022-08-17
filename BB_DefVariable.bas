Attribute VB_Name = "BB_DefVariable"
Option Explicit

'Page Number
Public PageNum As Integer

'Basic Visio Variables
Public vsoPage As Visio.Page
Public vsoPages As Visio.Pages
Public PageName As String
Public vsoConnectfromCell As Visio.Cell
Public vsoConnectToCell As Visio.Cell

Public vsoShapes As Visio.Shapes
Public vsoShapeNum As Integer
Public vsoShape As Visio.Shape
Public vsoConnect As Visio.Connect
Public vsoConnects As Visio.Connects

'Relation Table Variables
Public Relation(1 To 2000, 1 To relColMax) As Variant
Public RelationNo As Integer
Public RelationMaxNo As Integer

Public ConnectFromString As String
Public ConnectString  As String
Public ConnectStringStatus As Integer
Public ConnectStringTemp As String
Public ConnectToString As String

Public ConnectFromPort As String
Public ConnectToPort As String
Public ConnectFromName As String
Public ConnectToName As String
Public ConnectFromID As String
Public ConnectToID As String

'Basic Visio Shape Variables
Public ShapeDataList(1 To 5000, 1 To shdColMax) As Variant

'Max number of Antenna (ShuffleRelation)
Public AntCount As Integer

'Path Finding (ShuffleRelation)
Public LastPath As String
Public LastPath2 As String
Public UpPath As String

'Count Materials
Public MaterialList(1 To 1000, 1 To MatListColMax)

Public ShapeMaxRow As Integer
Public shdRowTemp As Integer
Public PageShapeIndex As Integer
Public MatInRow As Integer
Public BudInRow As Integer

Public LinkPath(1 To 1000, 1 To 3) As Variant
Public PathLoss(1 To 1000) As Double

'ErrorList Variables
Public ErrorExist As Boolean
Public ErrorNumber As Integer

Public SamePort(1 To 2000, 1 To 4) As Variant
Public PortRowCount As Integer

'In ShapeData
Public intRows As Integer
Public vsoCellValue As Variant

Public i As Integer
Public t As Integer

'Link Budget Label
Public NamingState As Integer

'Linkage
Public LinkNumOfRow As Integer
Public LinkRow As Integer
Public LinkSecRow As Integer
Public TestLinkRow As Integer
Public LinkageList(1 To 1000, 1 To lltColMax) As Variant
 
Public LinkageFromComp As Variant
Public LinkageFromPort As Variant
Public LinkageConnectors As Variant
Public LinkageToComp As Variant
Public LinkageToPort As Variant
Public j As Integer
Public k As Integer

'FloorList Array Variables
Public FloorList() As Variant
Public FloorRow As Integer
Public FloorMaxRow As Integer

'GetSector
Public SectorList() As Variant
Public SectorMaxNum As Integer
Public SectorName As Variant

'ExportExcelFile2
Public SectorRowCount As Integer

'BondOfMaterial
Public BondMatList() As Variant

'Autonumbering
Public AutoNum As Integer

'Userform
Public CheckButtonState As Integer
Public AutoPage As String

'RSRP and PathLoss
Public RSRP_output As Double
Public FSPL As Double
Public FSPL_lift As Double
Public FreqChoice As String
