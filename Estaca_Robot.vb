Imports Microsoft.VisualBasic

Public Class Class1

    ' This code creates a pile with elastic supports in Robot Structural Analysis

    Sub CreatePileWithElasticSupports()
        ' VARIABLES
        Dim RobApp As RobotApplication
        Dim RobDoc As RobotDocument
        Dim Nodes As RobotNodes
        Dim Bars As RobotBars
        Dim Supports As RobotSupports
        Dim Support As RobotNodeSupport
        Dim NodeIndex As Long
        Dim BarIndex As Long
        Dim PileLength As Double
        Dim NumSegments As Long
        Dim i As Long
        Dim SpringValue As Double
        Dim ExcelApp As Object
        Dim ExcelWorkbook As Object
        Dim ExcelSheet As Object
        Dim ExcelRow As Long

        ' SETTINGS
        PileLength = 30 ' Total length of the pile in meters
        NumSegments = PileLength ' 1m per segment
        ExcelRow = 2 ' Starting row in Excel (assuming row 1 is headers)

    ' INITIALIZE ROBOT
    Set RobApp = New RobotApplication
    RobApp.Project.New I_PT_SHELL
    Set RobDoc = RobApp.Project.Structure
    Set Nodes = RobDoc.Nodes
    Set Bars = RobDoc.Bars
    Set Supports = RobDoc.Supports
    
    ' OPEN EXCEL
    Set ExcelApp = GetObject(, "Excel.Application")
    Set ExcelWorkbook = ExcelApp.ActiveWorkbook
    Set ExcelSheet = ExcelWorkbook.Sheets(1)

    ' BEGIN MODELING
    ' Create nodes
    For i = 0 To NumSegments
            NodeIndex = i + 1
            Nodes.Create NodeIndex, 0, 0, -i ' Node spaced every 1m vertically down
        Next i

        ' Create bars
        For i = 1 To NumSegments
            BarIndex = i
            Bars.Create BarIndex, i, i + 1
    Next i

        ' Apply elastic supports
        For i = 1 To NumSegments + 1
            SpringValue = ExcelSheet.Cells(ExcelRow, 1).Value ' Read spring from column A
            ExcelRow = ExcelRow + 1
        
        Set Support = Supports.Create(i)
        Support.Constraints.X = csFree
            Support.Constraints.Y = csFree
            Support.Constraints.Z = csSpring ' Spring in Z
            Support.SpringZ = SpringValue
            Supports.Set i, Support
    Next i

        ' FINALIZE
        RobDoc.ViewMngr.Refresh
        MsgBox "Pile created with elastic supports!"

End Sub

End Class
